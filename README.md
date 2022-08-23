# coleta_dados_esocial

Na empresa em que atuo faço o controle das informações de SST (saúde e segurança do trabalho) de cerca de 500 empresas clientes, que devem ser enviadas para a plataforma do governo do eSocial mensalmente.

Por conta de uma procuração temos acesso as informações de SST das empresas clientes dentro da plataforma do eSocial do governo, utilizando o CPF do funcionário conseguimos acesso aos dados dos funcionários sem a necessidade de solicitar esses dados via email para as empresas.

Então, para atualizar o banco de dados do sistema que utilizamos para esses envios SOC, fiz uma automação para coleta de número de matrícula, situação do contrato e data de admissão dos funcionários a partir do CPF desses funcionários (que é cadastrado pelo cliente no SOC).

Para isso faço a extração de um planilha de informações dos funcionários cadastrados no sistema (chamada modelo1), leio essses dados com a biblioteca <b>openpyxl</b>, e utilizo a biblioteca <b>pyautoGUI</b> para automatizar a coleta desses dados na plataforma do governo do eSocial, após á coleta crio uma planilha com a relação dos dados coletados da plataforma do governo do eSocial e do SOC.

Após isso utilizo a modelo padrão de importação do sistema para atualizar os dados do funcionários.
