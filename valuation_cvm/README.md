# valuation_cvm
Obtenção de dados para valuation de empresas brasileiras por meio dos relatórios anuais e trimestrais da CVM em formato .xlsx. Cada empresa listada em valuation terá um arquivo de formato .xlsx salvo com os proncipais indicadores financeiros necessários para valuation da empresa.

# Utilização

Este guia descreve como criar um ambiente virtual Python para o projeto e posteriormente utilizar os códigos.

##Passo 1: Clonar o reposítório do GITHUB
Primeiro, clone o repositório usando o seguinte comando. Substitua `<url-do-repositorio>` pela URL do seu repositório GitHub:

```bash
git clone <url-do-repositorio>

## Passo 2: Navegar até o Diretório do Projeto
Após clonar o repositório, navegue até o diretório do projeto:

bash
cd nome_do_repositorio

## Passo 3: Criar o Ambiente Virtual
Crie um ambiente virtual com o seguinte comando. 

bash
python -m venv valuation_cvm_venv

## Passo 4: Ativar o Ambiente Virtual
Ative o ambiente virtual criado:

No Windows:
bash
nome_do_venv\Scripts\activate

No MacOS/Linux:
bash
source nome_do_venv/bin/activate

## Passo 5: Instalar as Dependências
Com o ambiente virtual ativado, instale as dependências (bibliotecas) do projeto usando o arquivo requirements.txt:

bash
pip install -r requirements.txt

## Passo 6: Verificar a Instalação
Para garantir que todas as bibliotecas foram instaladas corretamente, você pode listar os pacotes instalados:

bash
pip list
````
## Utilização dos Scripts
Utilize o Script "C:\Users\User\Desktop\valuation_cvm\valuation_cvm\Obtenção_de_Dados_CVM_B3_valuation.ipynb" para a obtenção dos dados financeiros a partir dos relatórios da CVM. É possivel selecionar o período de análise por meio da alteração dos nos inicial e final;
Após a obtenção dos relatórios, utilize o script "C:\Users\User\Desktop\valuation_cvm\valuation_cvm\valuation.ipynb" para filtrar e calcular os indicadores financeiros. A determinação da empresa que será obtida deve ser feita pela adição do nome e cnpj da empresa, seguindo o campo lista_empresas.

## Licença

Este projeto é licenciado sob a [Licença MIT](https://opensource.org/licenses/MIT). Consulte o arquivo `LICENSE` para mais detalhes.

## Contato

Para perguntas ou feedback, entre em contato comigo:

- **Nome:** Isa Lisboa
- **Email:** ilisboa@yahoo.com.br
- **GitHub:** [IsaGOLisboa]([https://github.com/IsaGOLisboa])

Sinta-se à vontade para abrir uma *issue* ou enviar um *pull request* se você encontrar algum problema ou tiver sugestões!





 



