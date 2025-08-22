<H1>Programa para Emissão Automática de NFe no Site Nota Carioca</H1>

O programa (complete_program.py) roda um frontend de fácil usabilidade conforme o print abaixo:

<img width="697" height="723" alt="final version" src="https://github.com/user-attachments/assets/82fdd7ac-3b5d-4954-ac2d-b1d4b23b6b6e" />
<p>

<p> 1. Para utilizar o bot, o primeiro passo é digitar o código, seguindo as instruções especificadas na tela abaixo.
Ao digitar o código, ele carrega na tela conforme as listas que estão dentro dos arquivos: item.py, servico.py, subitem.py </p>
<img width="492" height="179" alt="Screenshot 2025-08-22 at 02 03 53" src="https://github.com/user-attachments/assets/6df6429c-842c-4966-962b-028258c949ab" />

<p>

<p> 2. No segundo passo, ao clicar no botão de "fazer login" o chrome abre o site da nota carioca 
  (https://notacarioca.rio.gov.br/senhaweb/login.aspx) para que o usuário possa fazer login: </p>
<img width="995" height="605" alt="Screenshot 2025-08-22 at 02 25 51" src="https://github.com/user-attachments/assets/5ab0c82d-2ab3-469a-811b-a5b43fe32168" />

Obs: Eu optei por não implementar nenhum preenchimento automático com login/senha especialmente pra não ter que quebrar o captcha do site 
pois daria um trabalho desnecessário. Assim que essa tela é carregada é possível preencher tranquilamente com o login/senha antes de rodar a automação.

<p>
  
<p> 3. No terceiro passo, deve-se carregar um arquivo excel (.xlsx) que funcionará como uma base de dados, 
  em que cada linha contém todas as informações relativas a uma nota fiscal</p>
  
<img width="681" height="120" alt="Screenshot 2025-08-22 at 02 31 39" src="https://github.com/user-attachments/assets/3b156bf0-f6dd-4267-b1e4-b1e9e9303ec8" /><br>

A tabela em excel deve seguir a seguinte estrutura:
<img width="884" height="35" alt="Screenshot 2025-08-22 at 02 01 39" src="https://github.com/user-attachments/assets/77a1dd38-b297-4565-960f-522dccc37fe0" />

:warning: Pontos relevantes sobre a estrutura dos dados:

- Não alterar nenhum título de nenhuma coluna
- Os campos "Nome", "CPF" e "Preço do Produto" não podem ser deixados em branco
- Caso não tenha a informação preencher a célula com espaço
- O nome de pessoas que não são do RJ precisa estar completo
- É necessário ter uma coluna de status, que deve ser deixada em branco para que o bot preencha se foi emitido ou não.

<p> 4. No quarto passo, o usuário deve clicar na checkbox para validar que está com o navegador aberto para habilitar o botão para rodar a automação.</p>








