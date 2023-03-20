# RockNRoll Racing VB6 Modo Multiplayer
remake do jogo RockNRollRacing feito em VB6 (Visual Basic 6)
Projeto meramente educacional e não comercial , que demostra conceitos de isometria, sprites, colisão entre objetos no plano espacial. Possuindo apenas 1 carro e 3 pistas para demonstração . 
por motivo de direitos autorais , as trilhas sonoras não estão incluídas , somente os efeitos sonoros.
Código fonte em linguagem visual Basic 6.

# INSTALAR DLL binárias Windows
Será necessária a biblioteca OCX winsock e a DLL directx7 para vb6 
Após baixar este repositório,
abra o CMD em modo administrador.
vá para a pasta onde colocou os arquivos.
digite regsvr32 mswinsck.ocx
digite regsvr32 dx7vb.dll

# Executar EXE
Antes de rodar é preciso abrir o servidor
Execute RRRServer.exe
Será perguntando nome do servidor, número de voltas e número de participantes.

Execute RNRacing.exe

Para mudar o nome do jogador, pressione a tecla < ' > e digite name seu_apelido 

![rrr1](https://user-images.githubusercontent.com/25087767/226215653-15a9a186-e5d7-42fe-ad29-3946769a712d.png)

Novas cores (RGB)


![r2](https://user-images.githubusercontent.com/25087767/226216089-310ae567-6897-4540-8bc7-74508662c744.png)



Isometria e detecção de colisão 3D por vértices (vertex / vector collision)
Carros , objetos na pista, pista, rampas etc possuem posicionamento espacial , x, y, z. Assim , é possivel saber se o carro está dentro ou fora da pista, se está em uma rampa, subindo ou descendo e em quantos graus, se colidiu com outro objeto na pista, como óleo, se colidiu no ar com outro objeto.
![r3](https://user-images.githubusercontent.com/25087767/226217416-2921f6db-349d-48cd-88ab-f330a781fc03.png)
