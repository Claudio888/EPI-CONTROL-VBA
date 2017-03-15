# EPI-CONTROL-VBA
Objetivo: Realizar controle de EPI(Equipamento proteção individual) atraves de planilha do excel utilizando VBA.

Os usuarios são inseridos na planilha principal com as informações de email, nome, data do recebimento do EPI, o EPI, data de renovção,quantidade de dias de validade(oculto) e status. 

O código VBA ira analisar o valor do campo "status", se ele estiver com a palavra "Vencido" o script envia um email para o usuario indicando que ele deve realizar a troca de seu equipamento de proteção individual. 

Para definir "Vencido" ou "OK" é realizado uma comparação de datas, onde a celula I7 possui a data de Hoje-5(para que o email seja disparado com 5 dias de antecedencia ao vencimento), e então as celulas da coluna G (Status) comparam a data de hoje com a data de renovação(esta que é dada pela soma da data de recebimento e os dias de validade do EPI) e trocam o texto para "Vencido" ou "OK", que é comparado no VBA. 

Obs: Para funcionamento o aplicativo Outlook deve estar aberto e configurado no computador, pois o script ira utiliza-lo para envio de email, ou seja, o remetente sera a conta do outlook que estiver logada. 


