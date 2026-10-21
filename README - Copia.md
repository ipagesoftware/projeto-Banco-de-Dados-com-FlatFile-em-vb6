# Documentação da Tela — Cadastro de Produtos (Ver. 1.0)

Projeto desenvolvido em **Visual Basic 6.0 (VB6)** para demonstrar operações de CRUD (*Create, Read, Update, Delete*) utilizando arquivos de tamanho fixo como mecanismo de armazenamento de dados.

## Objetivo

Demonstrar técnicas de manipulação de arquivos em VB6, incluindo cadastro, consulta, alteração e exclusão de registros, sem a necessidade de um sistema gerenciador de banco de dados.

## Tecnologias

- Visual Basic 6.0
- Arquivos de acesso aleatório (*Random Access Files*)
- Registros de tamanho fixo

## Finalidade

Projeto educacional voltado ao estudo da programação em VB6, da manipulação de arquivos e da implementação de operações básicas de persistência de dados.

## Licença

Consulte os termos de licença definidos neste repositório.

---

> **Descrição funcional dos componentes de um CRUD de produtos.**

## TELA PRINCIPAL

![Tela_1](images/FrmMain6.png)

---

## 📌 Objetivo da Tela Principal
A tela **Principal (FrmMain)** possui os menus **Principal** e o menu **Ajuda**

Ela permite:
* Chamar o **Cadastro de Produtos**
* Encerrar a operação
* Acesso ao menu **Ajuda**

![Tela_2](images/FrmMain7.png)

![Tela_3](images/FrmMain5.png)

---

## Tela Cadastro de Produtos

![Tela_4](images/tela_cadproduto.png)

---

## Objetivo da Tela

A tela de **Cadastro de Produtos** é um CRUD (*Create, Read, Update e Delete*) destinado ao cadastro e à manutenção de produtos.

![Tela_6](images/FrmProduto2.png)

---

Ela permite:
* Localizar registros existentes
* Navegar entre os registros
* Incluir novos produtos
* Alterar dados
* Excluir registros
* Imprimir informações
* Encerrar a operação

---

## Fluxo de Utilização do CRUD

1. **Pesquisar:** O usuário informa o nome do produto no campo `NOME PRODUTO`.
2. **Localizar:** A lista apresenta os produtos compatíveis com o critério informado.
3. **Selecionar:** O usuário seleciona um produto na lista.
4. **Visualizar:** Os campos da área `DADOS` são preenchidos com as informações do registro selecionado.
5. **Navegar:** Os botões *Primeiro*, *Anterior*, *Próximo* e *Último* permitem percorrer os registros.
6. **Incluir:** O botão *Novo* inicia o cadastro de um produto.
7. **Alterar:** O botão *Alterar* permite modificar o produto selecionado.
8. **Excluir:** O botão *Excluir* remove o registro, preferencialmente com confirmação.
9. **Atualizar:** O botão *Atualizar/Recarregar* sincroniza a tela com a fonte de dados.
10. **Imprimir:** O botão *Imprimir* gera a saída impressa ou relatório correspondente.

---

## Estrutura Lógica Sugerida do CRUD

A tela é organizada conceitualmente em quatro grupos de operações:

* **CREATE (Criar):** Novo cadastro de produto.
* **READ (Ler):** Pesquisa, listagem, seleção e visualização dos produtos.
* **UPDATE (Atualizar):** Alteração dos dados do produto selecionado.
* **DELETE (Excluir):** Exclusão do produto selecionado.

---

## Observações Sobre a Implementação

* **Base de Dados em Arquivo:** A aplicação utiliza armazenamento via arquivo plano, conforme indicado na barra de status: `C:\ipagesoftware\Exemplos\FlatFile\database\produtos.dat`.
* **Mecanismo de Seleção:** A lista de produtos funciona como seleção dos registros, enquanto o painel `DADOS` apresenta os campos do item selecionado.
* **Tratamento no VB6 (ListBox ItemData):** Numa implementação em Visual Basic 6.0, o parâmetro `ItemData` da `ListBox` é utilizado para armazenar o código/ID do produto. Dessa forma, o texto
* exibido mostra o nome do produto enquanto `ItemData` mantém a identificação numérica para localização, alteração ou exclusão correta no banco.

---

### ACEITE E RESPONSABILIDADE DO USUÁRIO

Ao executar este instalador, o usuário declara estar ciente
de que o programa poderá realizar alterações no sistema
operacional Windows, criar, modificar ou remover arquivos e
diretórios, instalar componentes e executar processos com
privilégios administrativos.

A utilização deste instalador é realizada por conta e risco
do usuário.

O usuário é responsável por verificar previamente se possui
permissão para instalar, modificar ou remover os componentes
em seu computador.

---

## 📄 Licença e Uso

Este documento pode ser distribuído, modificado e compartilhado livremente para fins educacionais, estudos de arquitetura de software e análise de mercado de TI.
---

## Minhas Redes Sociais

Grupo no WhatsApp de Estudos: ** Código Limpo**
https://chat.whatsapp.com/HtA3mPmmB4RLw7tVJYVvL2

Grupo no WhatsApp da Ipage
https://chat.whatsapp.com/DPcG8meShJQCW3IgJLOYHZ

Grupo no Telegram
https://t.me/+RqTU5VvdkRvFCR36

Linkedin
http://www.linkedin.com/in/diogenes-dias-458a6a50

Instagram
https://www.instagram.com/ipage_software/?igsh=MWluYXhxcXE0cnE2cQ%3D%3D

Youtube
https://www.youtube.com/@Ipagesoftware

---

## Meus Produtos

APi para cálculo de rotas, consulta de CEP, consulta de CNPJ.
https://rapidapi.com/diogenes/api/ipage_cep/details

