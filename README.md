# ETIVisualizer
Um visualizador de arquivo .mdb para impressoras de Etiqueta.

## Objetivo
Criei esse programa para facilitar a minha vida e a dos amigos da Implantação, aonde trabalho. Pois um dos serviços envolve criar um Layout de etiqueta, 
porém de forma totalmente numérica e posicional, sem nenhum tipo de visualização. O que acarreta no alogamento do serviço e uso excessivo da impressora do cliente
para testes.

Pensando nisso criei um modo de pré-visualizar o layout por meio de uma imagem.

<img src= "https://user-images.githubusercontent.com/77119687/155725113-ff43406a-4057-4c9b-8eef-d1fb30503ce0.png" width= "240" height= "320"/>
Esse é um exemplo de um layout gerado, ele busca ser o mais fiél possivel a etiqueta real quando for impressa.

## Como baixar?
Fiz um vídeo curto explicando como fazer a instalação do programa. Basicamente entra na pasta [versions](https://github.com/EuReinoso/ETIVisualizer/tree/master/versions), baixa o mais recente e extrai. Mas caso tenha dúvidas assista o vídeo.

[Como Baixar (Youtube)](https://youtu.be/haRYrvIiZEc)

## Como usar?
Fiz um pequeno vídeo explicando como utilizar:

[Tutorial ETIvisualizer](https://youtu.be/CEe0qKWeOPg)

## Como funciona?
Explicação breve do funcionamento e suas dependencias.

### Tecnologias
- Python     - (Linguagem de Programação)
- Pyodbc     - (Biblioteca para leitura de banco de dados)
- Pandas     - (Biblioteca de manipulação de dados)
- Matplotlib - (Bliblioteca de visualização de dados em forma gráfica)

### Funcionamento
Primeiramente, o layout da etiqueta é armazenado em um arquivo ".mdb" - Microsoft Access Database - onde ficam armazenadas as tabelas posicionais dos campos das etiquetas.
Usando a bilbioteca Pyodbc é possivel ler essas tabelas e transformá-las em um arquivo ".csv", que é mais rápido e de mais facil manipulação.

Depois de converter as tabelas para um arquivo ".csv" é feita a leitura desse arquivo utilizando a biblioteca Pandas e os dados são transformados em uma imagem através da biblioteca Matplotlib.

### Curiosidades do funcionamento:
- #### Por que converter os arquivos ".mdb" para ".csv"?
R: Porque arquivos ".csv" são mais faceis e leves de manipular, e se adequam muito bem a biblioteca Pandas, a qual já sou familiarizado.

- #### Matplotlib
R: A biblioteca Matplotlib é na verdade utilizada para gerar gráficos e fazer análise de dados, porém ela conta com um recurso muito bom de posicionamento de dados
em um gráfico, que caiu como uma luva nesse caso.





