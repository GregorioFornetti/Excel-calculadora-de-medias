# Excel calculadora de medias
 Programa utilizando VBA para automatizar o cálculo de médias escolares no excel

# Tutorial
Aqui ensinarei a utilizar a calculadora de médias: como criar matérias, adicionar notas e visualizar os resultados

## Baixando o programa
Para baixar o programa, basta ir na [pagina principal do projeto](https://github.com/GregorioFornetti/Excel-calculadora-de-medias) e clicar no arquivo "Medias-excel.xlsm".

![Imagem com o arquivo necessário marcado](https://raw.githubusercontent.com/GregorioFornetti/Excel-calculadora-de-medias/main/tutorial/editadas/1.png)

Após fazer isso, clique em "Download" e o arquivo necessário para utilizar o programa começará a ser baixado.

![Imagem mostrando onde está o botão de download](https://raw.githubusercontent.com/GregorioFornetti/Excel-calculadora-de-medias/main/tutorial/editadas/2.png)

## Introdução

### Habilitar edição
Ao baixar o arquivo, abra-o e habilite a edição para que seja possível utilizar o programa.

![Imagem habilitando modo de edição]()

### Planilhas
No arquivo disponibilizado existem três planilhas

![Imagem das três abas de planilhas]()

1- Na primeira planilha, chamada "Menu", existem os botões para acionar os formulários para criação/exclusão de matérias ou notas

2- Todas as notas ficarão armazenadas na segunda planiha ("Notas"). Cada matéria terá suas notas separadas, com o valor afetado na nota final, o peso da nota e o valor da nota.

3- Todas as médias das matérias ficarão armazenadas na terceira planilha ("Médias"). Cada matéria terá sua média separada, calculando a média para cada tipo de nota e calculando a média final dependendo do peso de cada tipo de nota.

## Criando uma matéria

### Preenchendo o formulário

Para criar uma matéria, é necessário selecionar a planilha "Menu" e clicar no botão de criar matéria. Ao fazer isso, um formulário aparecerá.

![Formulário para criar uma matéria]()

No formulário, é preciso dar um nome para a matéria que será criada (não pode ser o mesmo nome de uma matéria previamente criada) e escolher a quantidade de tipos de notas que a matéria terá. Os tipos de notas são as divisões das notas que essa matéria terá para calcular a média final, por exemplo, em uma matéria, a média final pode ser calculada a partir de notas de provas, trabalhos e testes semanais. Cada tipo de nota precisa ter um nome único (a mesma matéria não pode ter tipos de notas repetidos) e um peso.
Segue abaixo um formulário de exemplo preenchido:

![Formulário de exemplo de criação de matéria preenchido]()

Nesse exemplo de formulário, estamos criando uma matéria chamada "Matemática" que possui 3 tipos de notas: provas (peso 2), trabalhos (peso 4) e testes (peso 3).
Com esses tipos de notas, a média final será calculada da seguinte forma:

![Imagem da fórmula para cálculo da média final da matéria de exemplo]()

Depois que preecher todo formulário, basta clicar no botão "Criar matéria" e se tudo tiver corretamente preechido a matéria será criada (se ocorrer um erro, uma mensagem aparecerá informando o que deve ser corrigido).

### Verificando criação da matéria

Para verificar se a matéria foi criada com sucesso, basta visitar as planilhas "Notas" e "Médias" e procurar campos com informações dessa matéria.
A matéria criada no exemplo ficou da seguinte forma:

Na planilha "Notas":
![Matéria exemplo na planilha notas]()

Na planilha "Médias":
![Matéria exemplo na planilha médias]()


## Adicionando notas

### Preenchendo o formulário
Para adicionar notas nas matérias criadas, basta ir na planilha "Menu" e clicar no botão de criação de notas. Ao clicar no botão, um formulário será aberto.

![Formulário de criação de notas]()

Nesse formulário, é necessário escolher a matéria na qual a nota será criada e o tipo de nota que a nota afetará. Após selecionar a matéria e o tipo de nota afetado, basta dar um nome único para essa nota (para que seja possível identifica-la unicamente), um peso e o valor da nota.
Logo abaixo está uma imagem com o formulário exemplo preenchido:

![Imagem do formulário de exemplo preenchido]()

Nesse formulário estamos criando uma nota para a matéria "Matemática" (a matéria criada no exemplo de criação de matérias) com o tipo de nota "Provas". Além disso, o nome da nota é "Prova 1", o peso é 1 e o valor da nota é 7.

### Verificando criação da nota

Após clicar em "Criar nota", é possível verificar na planilha de "Notas" que essa nota foi adicionada na matéria especificada,

![Imagem da planilha de notas após adicionar a nota]()

e que na planilha "Médias", a média já foi calculada automaticamente.

![Imagem da planilha de médias após adicionar a nota]()

## Apagando matérias e notas

### Apagando uma matéria
Para apagar uma matéria, basta clicar no botão "Excluir uma matéria" disponível na planilha "Menu". Ao clicar nesse botão, um outro formulário aparecerá, e nele basta escolher a matéria que você deseja excluir e clicar no botão "Excluir matéria".

![Imagem do formulário de exclusão de matéria]()

### Apagando uma nota
Para apagar uma nota, basta clicar no botão "Excluir uma nota" disponível na planilha "Menu". Ao clicar nesse botão, um formulário aparecerá, e nele basta escolher a matéria na qual a nota está e o nome da nota que você deseja excluir. Após selecionar as opções desejadas, clique em "Excluir nota" para efetuar a exclusão da nota.

![Imagem do formulário de exclusão de nota]()

## Avisos
1- Para evitar problemas com o programa, evite manipular as planilhas diretamente, apenas utilize os formulários disponíveis na planilha "Menu".

2- Não mude o nome das planilhas, isso causará problemas na execução dos formulários.
