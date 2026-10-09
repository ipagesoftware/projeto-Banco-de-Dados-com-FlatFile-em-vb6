# Documentação da Tela — Cadastro de Produtos (Ver. 1.0)

> **Descrição funcional dos componentes de um CRUD de produtos.**

## TELA PRINCIPAL

![Tela_1](images/FrmMain7.png)
---

## 📌 OBJETIVO DA TELA PRINCIPAL
A tela **Principal (FrmMain)** possui os menus **Principal** e o menu **Ajuda**

Ela permite:
* Chamar o **Cadastro de Produtos**
* Encerrar a operação
* Acesso ao menu **Ajuda**


## TELA CADASTRO DE PRODUTOS

![Tela_6](images/FrmProduto.png)

## 📌 OBJETIVO DA TELA

A tela de **Cadastro de Produtos** é um CRUD (*Create, Read, Update e Delete*) destinado ao cadastro e à manutenção de produtos.

Ela permite:
* Localizar registros existentes
* Navegar entre os registros
* Incluir novos produtos
* Alterar dados
* Excluir registros
* Imprimir informações
* Encerrar a operação

![Tela_6](images/FrmProduto2.png)

---

## 🔍 IDENTIFICAÇÃO GERAL DOS COMPONENTES

| Nº | Componente | Identificação | Descrição |
| :-: | :--- | :--- | :--- |
| **1** | Formulário principal | Janela `'CADASTRO PRODUTOS'` | Contêiner principal de toda a aplicação. Organiza menus, barra de ferramentas, área de pesquisa, dados do produto e informações de status[cite: 11]. |
| **2** | Menu Principal | `Principal` | Menu superior destinado às operações gerais do cadastro e às funções principais do módulo[cite: 11]. |
| **3** | Menu Navegação | `Navegação` | Menu relacionado à navegação e localização dos registros do cadastro[cite: 11]. |
| **4** | Área de orientação | Ícone de pesquisa + texto explicativo | Informa ao usuário que a caixa de texto deve ser utilizada para pesquisar e que a barra de ferramentas permite interagir com os dados[cite: 11]. |
| **5** | Barra de ferramentas | Toolbar horizontal | Concentra as operações de navegação, atualização, inclusão, alteração, exclusão, impressão e saída[cite: 11]. |
| **6** | Grupo de pesquisa | `PESQUISAR` | Área destinada à localização de produtos[cite: 11]. |
| **7** | Campo de pesquisa | `NOME PRODUTO` | Caixa de texto utilizada para digitar o nome ou parte do nome do produto a ser localizado[cite: 11]. |
| **8** | Lista de produtos | `ListBox` | Exibe os produtos encontrados. O usuário pode selecionar um item para visualizar seus dados no painel à direita[cite: 11]. |
| **9** | Contador de registros | `Total Registro(s): 1 de 25` | Indica a quantidade de registros disponíveis e a posição do registro atualmente selecionado[cite: 11]. |
| **10** | Grupo de dados | `DADOS` | Área que apresenta os campos do produto selecionado e permite sua manutenção[cite: 11]. |
| **11** | Código | Campo numérico | Identificador do produto. Normalmente corresponde à chave primária do registro[cite: 11]. |
| **12** | Nome | Campo de texto | Descrição/nome comercial do produto[cite: 11]. |
| **13** | Preço de compra | Campo monetário | Valor de aquisição ou custo unitário do produto[cite: 11]. |
| **14** | Preço de venda | Campo monetário | Valor utilizado para venda do produto[cite: 11]. |
| **15** | Quantidade | Campo numérico | Quantidade de unidades ou estoque associado ao produto[cite: 11]. |
| **16** | Total pago | Campo monetário | Valor total pago, normalmente relacionado à quantidade adquirida multiplicada pelo preço de compra ou ao valor registrado na operação[cite: 11]. |
| **17** | Preço tabela | Campo monetário | Preço de referência/tabela do produto[cite: 11]. |
| **18** | Último preço | Campo monetário | Último preço registrado para o produto[cite: 11]. |
| **19** | Barra de status | Caminho do arquivo | Exibe o local da base de dados utilizada pela aplicação (`produtos.dat`)[cite: 11]. |

---

## 🛠️ BARRA DE FERRAMENTAS — FUNÇÕES

| Ícone / Ação | Operação | Função |
| :-: | :--- | :--- |
| `|<` | **Primeiro registro** | Move o cursor para o primeiro produto do cadastro[cite: 11]. |
| `<` | **Registro anterior** | Volta um registro na sequência atual[cite: 11]. |
| `>` | **Próximo registro** | Avança um registro na sequência atual[cite: 11]. |
| `>|` | **Último registro** | Move o cursor para o último produto do cadastro[cite: 11]. |
| 🔄 | **Atualizar / Recarregar** | Atualiza os dados exibidos, relendo ou sincronizando os registros da fonte de dados[cite: 11]. |
| ➕ | **Novo / Incluir** | Inicia o cadastramento de um novo produto, limpando/preparando os campos para edição[cite: 11]. |
| ✏️ | **Alterar / Editar** | Coloca o produto atual em modo de edição para permitir alterações[cite: 11]. |
| ❌ | **Excluir** | Remove o produto selecionado, normalmente após confirmação do usuário[cite: 11]. |
| 🖨️ | **Imprimir** | Executa ou prepara a impressão dos dados/relatório relacionado ao cadastro[cite: 11]. |
| 🚪 | **Sair / Fechar** | Encerra a tela atual ou retorna ao módulo anterior[cite: 11]. |

---

## 🔄 FLUXO DE UTILIZAÇÃO DO CRUD

1. **Pesquisar:** O usuário informa o nome do produto no campo `NOME PRODUTO`[cite: 11].
2. **Localizar:** A lista apresenta os produtos compatíveis com o critério informado[cite: 11].
3. **Selecionar:** O usuário seleciona um produto na lista[cite: 11].
4. **Visualizar:** Os campos da área `DADOS` são preenchidos com as informações do registro selecionado[cite: 11].
5. **Navegar:** Os botões *Primeiro*, *Anterior*, *Próximo* e *Último* permitem percorrer os registros[cite: 11].
6. **Incluir:** O botão *Novo* inicia o cadastro de um produto[cite: 11].
7. **Alterar:** O botão *Alterar* permite modificar o produto selecionado[cite: 11].
8. **Excluir:** O botão *Excluir* remove o registro, preferencialmente com confirmação[cite: 11].
9. **Atualizar:** O botão *Atualizar/Recarregar* sincroniza a tela com a fonte de dados[cite: 11].
10. **Imprimir:** O botão *Imprimir* gera a saída impressa ou relatório correspondente[cite: 11].

---

## 🧱 ESTRUTURA LÓGICA SUGERIDA DO CRUD

A tela é organizada conceitualmente em quatro grupos de operações[cite: 11]:

* **CREATE (Criar):** Novo cadastro de produto[cite: 11].
* **READ (Ler):** Pesquisa, listagem, seleção e visualização dos produtos[cite: 11].
* **UPDATE (Atualizar):** Alteração dos dados do produto selecionado[cite: 11].
* **DELETE (Excluir):** Exclusão do produto selecionado[cite: 11].

---

## ⚙️ OBSERVAÇÕES SOBRE A IMPLEMENTAÇÃO

* **Base de Dados em Arquivo:** A aplicação utiliza armazenamento via arquivo plano, conforme indicado na barra de status: `C:\ipagesoftware\Exemplos\FlatFile\database\produtos.dat`[cite: 11].
* **Mecanismo de Seleção:** A lista de produtos funciona como seleção dos registros, enquanto o painel `DADOS` apresenta os campos do item selecionado[cite: 11].
* **Tratamento no VB6 (ListBox ItemData):** Numa implementação em Visual Basic 6.0, o parâmetro `ItemData` da `ListBox` é utilizado para armazenar o código/ID do produto[cite: 11]. Dessa forma, o texto exibido mostra o nome do produto enquanto `ItemData` mantém a identificação numérica para localização, alteração ou exclusão correta no banco[cite: 11].

---

## 🏗️ VISÃO GERAL DA ARQUITETURA DE TELAS

O módulo é composto por uma **tela principal de navegação e consulta (`FrmProduto`)**, duas **janelas modais de operação (`FrmInsert` e `FrmEdit`)**, uma **caixa de diálogo para confirmação de exclusão** e um **gerador/preview de relatório**[cite: 11].
