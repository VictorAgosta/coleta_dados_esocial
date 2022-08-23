# coleta_dados_esocial

Na empresa em que atuo faço o controle das informações de SST (saúde e segurança do trabalho) de cerca de 500 empresas clientes, que devem ser enviadas para a plataforma do governo do eSocial mensalmente.

Como por conta de uma procuração temos acesso as informações de SST das empresas clientes, utilizando o CPF do funcionário conseguimos acesso aos dados dos funcionários sem a necessidade de solicitar esses dados via email para as empresas.

Então, para atualizar o banco de dados do sistema que utilizamos para esses envios SOC, fiz uma automação para coleta de número de matrícula, situação do contrato e data de admissão dos funcionários a partir do CPF desses funcionários (que é cadastrado pelo cliente no SOC).

Para isso faço a extração de um planilha de informações dos funcionários cadastrados no sistema (chamada modelo1), leio essses dados com 
