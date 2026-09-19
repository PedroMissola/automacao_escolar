# Política de Segurança

A segurança do **Motor de Automação Escolar** e a integridade dos dados processados são prioridades. Agradecemos a contribuição da comunidade em ajudar a manter este projeto seguro.

## Reportando uma Vulnerabilidade

Se você descobriu uma vulnerabilidade de segurança neste projeto, pedimos que **não** a divulgue publicamente (criando Issues ou Pull Requests abertos) até que possamos corrigi-la.

**Para reportar uma vulnerabilidade, entre em contato através dos seguintes canais:**

- **E-mail:** [missolapedro@gmail.com]
- **Forma alternativa:** Abra uma Issue **privada** (se o repositório estiver configurado para isso) ou entre em contato diretamente com o gestor do projeto (Pedro).

## O que esperar ao reportar

1. **Confirmação:** Responderemos à sua notificação em até **48 horas** para confirmar o recebimento da sua descoberta.
2. **Análise:** Investigaremos o relatório para validar a existência e o impacto da vulnerabilidade.
3. **Correção:** Trabalharemos em uma correção assim que a vulnerabilidade for confirmada. O prazo para resolução dependerá da complexidade do problema.
4. **Divulgação:** Após a correção ser implantada, publicaremos um agradecimento (com o seu consentimento) e atualizaremos a documentação, se necessário.

## Práticas de Segurança Adotadas no Projeto

- **Isolamento de Credenciais:** As chaves de API (como a do Resend) e endereços de e-mail são gerenciados exclusivamente via variáveis de ambiente (`.env`) e não são versionados no repositório.
- **Prevenção de Corrupção de Dados:** O sistema utiliza arquivos de trava (*lockfiles*) para impedir execuções simultâneas do script, garantindo que o banco de dados e as planilhas não sejam corrompidos em caso de múltiplas chamadas.
- **Processamento Local:** O banco de dados (SQLite) e a lógica de validação rodam de forma local e offline. Isso evita a exposição de informações acadêmicas em redes públicas ou servidores web.
- **Integridade e Rastreabilidade:** O código utiliza módulos nativos de criptografia para gerar hashes (SHA1) únicos para cada atividade processada, o que impede a duplicação de informações no banco de dados.

## Escopo

Esta política de segurança se aplica estritamente ao código-fonte deste repositório. Não cobre vulnerabilidades em serviços de terceiros utilizados (como a API da Resend) ou no próprio ambiente Node.js, que possuem suas próprias políticas.

---

*Agradecemos por nos ajudar a manter o Motor de Automação Escolar seguro!*
