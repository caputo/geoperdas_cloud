# geoperdas_cloud
 Extração do software GeoPerdas da ANEEL para execução em multiplos containers linux.

 Com essa solução é possivel realizar o escalomento do calculo em um ambiente cloud ampliando o numero de replicas do container de calculo, permitindo assim a execução de um calculo de forma mais rapida e permitindo o desprovisionamento de recursos de maquina apos o termino do calculo. 

 A solicitacao do calculo é feito através de um portal WEB, os alimentadores são enviado para uma fila e os motores de calculo são reposavel por ler a fila e executar o calculo de cada alimentador individualmente. 


A solução é composta pelos sequintes componentes. 

# Geoperdas Console
Esse é o resposavel por toda execução do calculo. O código foi extraido do programa ProgGeoPerdas para Desktop e removido toda a interface para Desktop.

OpenDSS. 
Para que seja possivel executar em um container linux foi retirado a dependecia via interface COM e utilizando a dll do OpenDSS compativel. 

# WebAPI 
Portal responsavel por criar a solicitação indicano os parametros de calculo e enviando os alimentadores para a fila. 

# RabbitMQ
Gerenciador de filas responsavel por receber as solicitacoes de calculo de alimentadores e prover para os motores de calculo. 

# SQL Server
O banco de dados pode ser executado via container junto com os demais ou estar localizado em um outro servidor dedicado (recomendado)


# MODO DE USO


Single Server - Docker Compose

- Necessário ter docker instalado na maquina HOST
- Verifique o arquivo docker-compose.yml e valide quais containers são necessários e quantas réplicas de motores sao necessarias para a execucao
- Executar docker-compose up na pasta raiz do projeto. 


TODO
Via Kubernets 
Azure Cloude
