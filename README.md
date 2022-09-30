# ManipulandoXLSX
<p>Projeto cuja função é recuperar dados de um arquivo txt e criar e inserir os dados recuperados em um arquivo xlsx</p>
<div align="center">
<img src="https://user-images.githubusercontent.com/109561598/193161561-fceba7cf-38cb-460d-ab44-1e20c58b4a06.png" height="200px"/>
</div>

## Resumo:
<p align="justify">
    Por questões práticas e tecnológicas diversas pessoas, principalmente em ambiente organizacional optam por fazerem relatórios e planilhas digitalizados por computadores para assim acessá-los e apresentá-los de forma mais eficiente.<br />
    Assim sendo, a fim de facilitar a vida de um colaborador da área da psicologia que tem a função de fazer um relatório em um arquivo de texto sobre seus pacientes e, após, passar os dados desse relatório em forma de planilha. Desenvolvi um programa que acessa este arquivo na extensão ".txt" e recupera as informações deste arquivo (nome, quantidade de sessões no mês e convênio), calcula o total a receber de cada cliente levando em consideração a quantidade de consultas com o quanto o convênio do respectivo cliente paga.    
</p>

## Tópicos: 
:small_blue_diamond: [Sobre o projeto](#sobre-o-projeto)

:small_blue_diamond: [Funções do script](#funções-do-script)

:small_blue_diamond: [Requisitos para rodar o código](#requisitos-para-rodar-o-código)

:small_blue_diamond: [Como rodar o código](#como-rodar-o-código)

## Sobre o Projeto:
<p align="justify">O programa foi escrito em Python com o auxílio da biblioteca openpyxl, usada para que o Python possa interagir com arquivos gerados pela ferramenta de edição e criação de planilhas Microsoft Excel (arquivos de extensão ".xls" e ".xlsx"). O diretório conta com o script app.py e duas pastas, documentos (onde os documentos de extensão ".txt" ficam) e planilhas (onde o programa app.py salva as planilhas criadas)

</p>

## Funções do script:
<ul>
    <li align="justify">
        <h4>main():</h4> Função principal, faz um loop enquanto interage com o usuário para receber o nome do relatório para recuperar os dados e o mês referente à planilha além de chamar as demais funções para criar, inserir os dados e formatar a planilha.</li>
    <li align="justify">
        <h4>txt(arq):</h4> Função que recebe o nome do arquivo de texto inserido pelo usuário; verifica se foi adicionado a extensão, caso não, adiciona a extenção".txt".</li>
    <li align="justify">
        <h4>usaDoc(documento):</h4> Função que recebe o nome do documento; abre o documento; separa os registros; verifica para cada registro, para cada linha do registro se possui algum dos dados necessários, caso a linha tenha a função chamará outra função que garimpa apenas o valor necessário; retorna um dicionário contendo todos dados dos registros.</li>
    <li align="justify">
        <h4>precoConvenio(convenio,qtd):</h4> Função que recebe o nome do convênio do registro e a quantidade de seções; verifica o convênio e retorna o resultado do preço por seção multiplicado pela quantidade.</li>
    <li align="justify">
        <h4>pegaQtd(frase):</h4> Função que recebe a linha onde o valor referente à quantidade se encontra e retorna apenas este valor.</li>
    <li align="justify">
        <h4>pegaNome(frase):</h4> Função que recebe a linha onde o valor referente ao nome se encontra e retorna apenas este valor.</li>
    <li align="justify">
        <h4>repetido(lista,nome):</h4> Função que recebe uma lista de dicionários e um nome; a função irá verificar se o nome pertence a algum dicionário da lista e, caso sim: irá retornar uma lista com o booleano verdadeiro e o índice do dicionário onde está o nome, caso não: será retornado uma lista com falso e o número -1.</li>
    <li align="justify"> 
        <h4>criaPlan(arq,mes):</h4> Função que recebe o nome do arquivo texto e o mês referente à planilha; cria a planilha; chama as funções que adicionam e estiliza os registros e salva o arquivo com extensão ".xlsx".</li>
    <li align="justify">
        <h4>addPessoas</h4>(pessoas, plan, linhaInicial):</b> Função que recebe uma lista de dicionários que contêm os dados das pessoas, a planilha em que serão inseridos os dados e a linha de onde o programa iniciará a adicionar; adiciona os registros e seus valores na planilha.</li>
    <li align="justify">
        <h4>estiliza(planilha,celulas,estilo):</h4> Função que recebe a planilha que irá modificar, uma lista com as células e um inteiro que indica o tipo de estilo; caso estilo seja igual a 1, a função irá estilizar as células como células de título; caso estilo seja igual a 2, as células serão formatadas como células de valores.</li>
</ul>

## Requisitos para rodar o código:

<ul>
    <li>Linguagem Python</li>
    <li>Biblioteca openpyxl</li>
</ul>
<div style="display:flex">
    <img src="https://raw.githubusercontent.com/github/explore/80688e429a7d4ef2fca1e82350fe8e3517d3494d/topics/python/python.png" height="60px" width="60px"/>
    <img src="https://user-images.githubusercontent.com/109561598/193160002-cd3e64d1-0076-4403-ae89-f870db7d5a43.png" height="50px" />
</div>
<p align="justify">Recomendo que use uma versão recente do python para evitar erros. A versão que usei foi a 3.8</p>
<p>Para instalar o Python em seu computador,acesse o link abaixo para abrir o site oficial da linguagem onde pode-se consultar informações sobre a instalação:</p>

```
https://www.python.org
```
<p align="justify">Para instalar o openpyxl cole o seguinte código na linha de comando do seu computador:</p>

```
pip install openpyxl
```
<p>Para mais informações sobre a biblioteca acesse a documentação disponível em:</p>

```
https://openpyxl.readthedocs.io/
```

## Como rodar o código:
<ol>
<li align="justify">
    Selecione a pasta que deseja baixar o projeto, abra a interface de linha de comando do Git e cole o seguinte código para clonar o projeto:

```
git clone https://github.com/KevinFGR/ManipulandoXLSX.git 
```
</li>
<li align="justify">
    Abra a interface de linha de comando do seu computador na pasta manipulandoxlsx e insira o seguinte comando para executar o programa:

```
py app.py
```
</li>
<li>Interaja com o programa respondendo o mês referente e o nome do arquivo de texto</li>
<li>Após o programa informar que a planilha foi criada você pode encontrá la na pasta planilhas</li>
</ol>

