# controle-solicitacao-compra-recebimento
## Controle para visualização do status do processo de compra das matérias primas

### Problemática:
Desenvolver uma ferramenta capaz de retornar o status da compra de uma matéria prima solicitada para um projeto


### View em T-SQL: (view-materia-prima.sql)
Esta visualização é capaz de percorrer todo o processo de aquisição de matéria-prima do ERP da empresa, inclusive identificando o projeto atrelado ao material e o cliente envolvido.
Por este motivo, o código possui joins para a parte de orçamento e serviços, bem como para tabelas de solicitação, cotação, ordens de compra e entrada de material no estoque.
Por fim, baseado nos campos extraídos o material é classificado de acordo com a localização e/ou data prevista de entrega.


### Script Python: (script-automacao.ipynb)
A finalidade do script é automatizar o processo de sincronização do banco de dados. Com a biblioteca win32com o código abre o arquivo, sincroniza todas as informações com o banco de dados, salva e encerra a instância Excel.


### Script VBA: (vba-filtrar-por-servico.cls)
É utilizada dentro do arquivo excel com conexão ao banco de dados. Sua finalidade é facilitar a busca do serviço digitando o número em uma célula.
O valor preenchido pelo usuário é utilizado para filtrar a tabela dinâmica e retornar apenas os materiais do projeto desejado.