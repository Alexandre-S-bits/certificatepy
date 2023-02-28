# Certificatepy
## _Gerador de Certificados_  
&nbsp;
## Sobre
___
Certificatepy é um gerador de certificados escrito em python com interface utilizando tkinter, o objetivo é gerar certificados em massa com poucos cliques. 
- Os dados são obtidos com o preenchimento da planilha do Excel
- Pode-se separar em diferentes guias da planilha informações para ajudar no controle dos certificados, e no momento de gerar basta escolher a guia desejada
- Com a guia *'coordenadas'* é possível alterar a posição de preenchimento das informações, assim é possível gerar certificados em diferentes layouts
- As fontes padrões podem ser alteradas para combinarem com o layout do certificado  
&nbsp;
##  Como usar
___
1. Insira as informações na planilha *'informacoes.xlsx'*, insira os dados que serão preenchidos no certificado e as coordenadas, e SALVE a planilha, pois os dados não serão capturados caso não salve.
2. No programa clique em **Selecionar planilha** e escolha a planilha *'informacoes.xlsx'*
3. No menu abaixo **Nome da guia** selecione a guia de onde estará capturando as informações
4. Clique em **Escolher imagem** e escolha o template base do certificado (todos os certificados serão gerados a partir deste arquivo). Marque a caixa de texto **Padrão** caso exista um arquivo com nome *'certificate.png'* (O formato da imagem deve ser PNG) na mesma pasta do arquivo **certificatepy** ou do **executável**
5. Selecione a Fonte das letras do arquivo, é possível mudar o tipo de fonte, para isso leia nas Opções Extras
6. Pressione o botão **Gerar Certificados**, o local de salvamento será na mesma pasta que está sendo executado a aplicação  
&nbsp;
## Informações Úteis
___
- É possível criar várias guias com  dados de preenchimento no certificado, por exemplo, guias separando por turma ou período.
- Para mudar o tipo de fonte, mova o arquivo da fonte para a pasta *'fonts'* e substitua por uma nova. 
>Obs: É necessário que as fontes tenham o formato .ttf ou .otf, e que elas sejam nomeadas EXATAMENTE com o mesmo nome de alguma fonte da pasta.
- Altere as coordenadas dos dados na guia 'coordenadas' da planilha *'informacoes.xlsx'*, caso precise da coordenada, abra a imagem do certificado e utilize o paint ou algum programa de edição de imagens.  
&nbsp;
## Bibliotecas requeridas
Para execução do código é necessário as bibliotecas abaixo:
| Bibliotecas | Documentação |
| ------ | ------ |
| openpyxl | https://openpyxl.readthedocs.io/en/stable |
| pillow | https://pillow.readthedocs.io/en/stable/ |
| tkinter | https://docs.python.org/3/library/tkinter.html |

&nbsp;
## Criando um arquivo executável para Windows
___
Para criar um arquivo **.exe** instale a biblioteca pyinstaller:
```sh
pip install -U pyinstaller
```

Para salvar o arquivo como apenas um arquivo execute o comando abaixo:
```sh
pyinstaller certificatepy.py –windowed
```
Se quiser consulte a documentação completa do [pyinstaller](https://pyinstaller.org/en/stable/)
