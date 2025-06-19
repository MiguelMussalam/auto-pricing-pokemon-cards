> **Aviso de uso responsável**  
> Este projeto foi desenvolvido **exclusivamente para fins educacionais** e para aprimoramento pessoal em automação com Python e manipulação de planilhas.  
> **Não é recomendado o uso frequente ou em larga escala.**  
> O código **não implementa medidas para contornar limites do site da Liga Pokémon**, como o número máximo de itens no carrinho.  
> Se for utilizá-lo, faça com moderação e respeito à plataforma.

# Auto Pricing Pokémon Cards
Automatize a coleta de preços de cartas Pokémon usando Selenium com suporte a planilhas do Excel.

## Sobre o projeto

Este projeto utiliza Python + Selenium para:
- Ler uma planilha `.xlsx` com informações de cartas Pokémon (nome, número, condição)
- Acessar o site da [Liga Pokémon](https://www.ligapokemon.com.br)
- Coletar automaticamente o preço da carta com base na condição
- Inserir o preço e a data diretamente no Excel
- Calcular o valor do album e editar data de modificação

## Estrutura da planilha esperada

A planilha Excel necessita ter o nome: (`album.xlsx`), contendo as seguintes colunas:

| Nome      | Número     | Condição | Preço | Ilustração | Total | Data        |
|-----------|------------|----------|--------|-------------|--------|-------------|
| Ádamo     | GG56/GG70  | NM       |        |             |        |             |
| Arceus V  | GG70/GG70  | NM       |        |             |        |             |
| ...       | ...        | ...      | ...    |    ...      | ...    | ...         |

- **Nome** e **Número** _precisam_ estar preenchidas, caso não estejam, a carta é pulada.
- O nome e número da carta precisa ser _exatamente_ o mesmo nome apresentado no site.
- **Condição** aceita valores: `M`, `NM`, `SP`, `MP`, `HP`, `D`
- **exemplo:** A carta no site é: Pikachu (001/007), portanto, **nome: **`Pikachu` N**úmero:** `(001/007)`
- Uma planilha com a formatação já pronta para o script está no repositório com o nome de `album.xlsx`

- **OBS:** A coluna **Ilustração** está reservada para carregar imagens da carta automaticamente (em desenvolvimento).

## Requisitos

- Python 3.10+

## Instalação

```bash
git clone https://github.com/MiguelMussalam/auto-pricing-pokemon-cards.git
cd auto-pricing-pokemon-cards
python -m venv .venv
source .venv/bin/activate  # ou .venv\Scripts\activate no Windows
pip install -r requirements.txt
