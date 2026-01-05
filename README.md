# Gestao_Despesas

Repositório de códigos utilizados nas **Solicitações – Gestão de Despesas Corporativas (Vexpenses e Onfly)**.

---

## Automação de Aprovação de Solicitações de Aplicativos (APPs)

**Data da atualização:** 20/10/2025  
**Responsável:** Lucas Garcia — Analista de Melhoria Contínua  
**Tipo de atualização:** Automação de processos / Substituição de plataforma / Padronização de comunicação  
**Status:** Implementado  

---

## Contexto

O processo de solicitação e aprovação de aplicativos (APPs) era anteriormente executado por meio de um fluxo desenvolvido no **Power Automate**, o qual passou a apresentar recorrentes falhas operacionais, incluindo erros de execução, perda de eventos de aprovação, instabilidade na comunicação por e-mail e baixa rastreabilidade das decisões tomadas.

Além da instabilidade técnica, o fluxo anterior dificultava ajustes evolutivos, dependia de múltiplos conectores externos e não oferecia flexibilidade suficiente para personalização da experiência do aprovador e do solicitante. Esse cenário gerava retrabalho, insegurança quanto ao registro das decisões e impacto direto na confiabilidade do processo de Qualidade.

Diante dessas limitações, tornou-se necessária a **reconstrução integral do fluxo**, adotando uma solução mais estável, controlável e aderente ao ecossistema Google já utilizado pela organização.

---

## Descrição da atualização

Foi desenvolvida e implementada uma nova automação utilizando **Google Forms, Google Sheets, Google Apps Script e Gmail**, substituindo completamente o fluxo anterior do Power Automate.

A solução foi desenhada para ser **robusta, auditável e escalável**, centralizando dados, decisões e comunicações em um único ecossistema, com controle total do código e das regras de negócio.

---

### 1. Coleta estruturada das solicitações

- As solicitações são registradas via **Google Forms**, garantindo padronização das informações desde a origem.  
- Cada submissão gera automaticamente um registro na planilha base (**“Respostas ao formulário 1”**).  
- Um **ID numérico sequencial** é gerado automaticamente para cada solicitação, assegurando rastreabilidade única e permanente.

---

### 2. Envio automático de e-mail de aprovação

- A cada nova submissão, é enviado um **e-mail HTML personalizado** ao gestor aprovador.  
- O e-mail contém:
  - **Identidade visual da empresa** (logotipo incorporado)  
  - **Dados da solicitação** (ID, data/hora, solicitante, aplicativo, descrição e observações)  
  - **Botões de ação direta** para **Aprovar** ou **Reprovar**  
  - **Link direto** para a planilha de acompanhamento  

---

### 3. Telas de aprovação personalizadas (Web App)

- Os botões de aprovação e reprovação direcionam para um **Web App desenvolvido em Google Apps Script**.  
- As telas são **totalmente personalizadas**, com layout responsivo e identidade visual institucional.  
- O sistema valida parâmetros (**ação** e **linha**) antes de registrar qualquer decisão.  
- Após a decisão, o aprovador visualiza uma **tela de confirmação**, contendo:
  - **Status aplicado**  
  - **Data e hora do registro**  
  - **Link para consulta na planilha**  

Mesmo que o Google exiba alertas de verificação de domínio, a decisão é corretamente registrada, eliminando riscos operacionais.

---

### 4. Registro automático da decisão

- A planilha é atualizada automaticamente com:
  - **Status da solicitação** (Aprovado / Reprovado)  
  - **Data e hora da decisão**  
  - **Comentários do gestor**, quando existentes  
- Todas as alterações ficam **centralizadas e auditáveis** em um único repositório.

---

## Comunicação e usabilidade

### E-mails personalizados e padronizados

- Todos os e-mails são enviados em **HTML**, com linguagem institucional e padronizada.  
- O fluxo diferencia:
  - E-mail de **solicitação de aprovação**  
  - E-mails de **reenvio automático ou manual**  
  - E-mail de **confirmação da decisão**  

---

### Confirmação automática ao solicitante

- Após a aprovação ou reprovação, o solicitante recebe automaticamente um **e-mail de confirmação**, contendo:
  - **Status final da solicitação**  
  - **Dados completos do pedido**  
  - **Data e hora da decisão**  
  - **Comentário do gestor**, quando informado  

Esse retorno elimina dúvidas, reduz contatos manuais e aumenta a **transparência do processo**.

---

### Reenvio controlado de solicitações

- O sistema permite:
  - **Reenvio automático** de solicitações pendentes  
  - **Reenvio manual por ID**, via menu personalizado na planilha  
- Solicitações já **aprovadas ou reprovadas** são automaticamente bloqueadas para reenvio indevido.

---

### Coleta de comentários via e-mail

- O fluxo realiza a **leitura automática** das respostas do aprovador no Gmail.  
- Apenas o texto digitado pelo gestor é considerado, desconsiderando:
  - Assinaturas  
  - Citações  
  - Respostas automáticas  
- O comentário é registrado diretamente na planilha, vinculado ao **ID da solicitação**.

---

## Impactos e benefícios

A implementação desta automação trouxe ganhos diretos e estruturais para o processo de Qualidade, incluindo:

- Eliminação d
