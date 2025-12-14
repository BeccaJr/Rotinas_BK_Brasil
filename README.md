# 📊 Rotinas Operacionais BK Brasil (VB.NET)

Automação de rotinas operacionais críticas utilizadas no reporte diário dos restaurantes **Burger King** e **Popeyes** no Brasil.

Este projeto foi desenvolvido em **2020**, durante minha atuação como **Assistente de CCO no Burger King (atual ZAMP)**, com o objetivo de eliminar processos manuais, reduzir erros humanos e aumentar a confiabilidade dos indicadores operacionais da companhia.

---

## 🎯 Contexto de Negócio

Na época, não existia um banco de dados centralizado.  
Os principais indicadores operacionais da empresa eram gerados a partir de:

- Múltiplos relatórios extraídos de sistemas distintos  
- Consolidação manual em Excel  
- Uso intensivo de fórmulas, cópia/cola e ajustes manuais  

Essa rotina:
- Começava diariamente às **4h da manhã**
- Levava até **4 horas** para ser concluída
- Era altamente suscetível a **erros humanos**
- Impactava diretamente KPIs enviados **do CEO até os gerentes das lojas**

Pequenas inconsistências acumuladas ao longo do mês afastavam o time do KPI real e dificultavam tomadas de decisão.

---

## 💡 Proposta da Solução

Identificando que o problema era **estrutural**, desenvolvi uma aplicação em **VB.NET** que automatiza todo o processo de consolidação dos relatórios após o download das bases.

A solução foi pensada para:
- Padronizar processos inexistentes até então
- Eliminar interferência manual
- Garantir consistência e rastreabilidade dos dados
- Ser simples o suficiente para qualquer analista operar

---

## 🛠️ Tecnologias Utilizadas

- **VB.NET**
- **Windows Forms**
- **Microsoft Excel (automação)**
- **VBA (validações internas nos relatórios)**
- **Tabelas e Gráficos Dinâmicos**
- **Formatação Condicional**
- **Validações de estrutura e colunas**

---

## ⚙️ Funcionalidades Principais

- Seleção guiada dos arquivos de entrada  
- Validação automática de:
  - Arquivos ausentes
  - Seleção incorreta
  - Estrutura e padrão de colunas
- Consolidação automática das bases
- Geração de relatórios operacionais:
  - MTD (Month to Date)
  - D-1
- Eliminação total de edição manual
- Interface simples e orientada ao fluxo do usuário

---

## 🧠 Principais Desafios Resolvidos

- Padronização de arquivos vindos de **sistemas diferentes**
- Tratamento de problemas de **encoding e formatação**
- Performance no processamento de grandes volumes de dados
- Redução de erros causados por fórmulas inconsistentes
- Simplificação do processo para escalabilidade do time

---

## 📊 Resultados Alcançados

- ❌ Erros humanos reduzidos a **zero**
- ⏱️ Tempo médio diário reduzido de ~4h para **2h30**
- 📈 KPIs mais confiáveis
- 🧠 Mais tempo dedicado à análise, menos à operação
- 🧩 Processo reutilizável por novos integrantes do time

---

## 📸 Interface da Aplicação

A aplicação possui uma interface simples e orientada à execução do processo, com:
- Menu principal de seleção
- Painéis operacionais
- Validações visuais de erros
- Mensagens claras para o usuário

---

## 🚀 Como Executar

1. Clone o repositório:
```bash
```git clone https://github.com/BeccaJr/Rotinas_BK_Brasil.git```

2. Abra a solução no Visual Studio

3. Compile o projeto

4. Execute o aplicativo

5. Selecione os arquivos conforme solicitado pela interface

---

## 📌 Observações Importantes

- Este projeto reflete um contexto real de negócio da época
- Não utiliza banco de dados, pois a infraestrutura ainda não existia
- O foco é automação, padronização e confiabilidade
- Código disponibilizado para fins educacionais e de portfólio

---

## 🎥 Demonstração

Em breve: vídeo demonstrando o funcionamento completo da aplicação.

---

## 👤 Autor

Desenvolvido por BeccaJr

📎 LinkedIn: https://www.linkedin.com/in/beccajr/
📂 GitHub: https://github.com/BeccaJr

---

## 🧠 Filosofia

“Sou um profissional preguiçoso — do tipo que prefere automatizar hoje para não repetir o mesmo trabalho amanhã.”
