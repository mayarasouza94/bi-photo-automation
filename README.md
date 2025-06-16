# Automação de Fotos para Power BI

Este projeto automatiza a extração, classificação e formatação de imagens provenientes de apresentações em PowerPoint (relatórios semanais de obra), estruturando-as em uma base compatível com o Power BI. As imagens são convertidas em base64 e enriquecidas com dados extraídos automaticamente, como data, pavimento, serviço, sublocalização e categoria.

## Problema Resolvido

Em muitos projetos de engenharia, fotos de progresso são arquivadas manualmente, dificultando padronização, rastreabilidade e visualização integrada em BI. Este script elimina esse gargalo, permitindo que relatórios fotográficos sejam transformados em dados navegáveis e visuais.

## Funcionalidades

- Extração automática de imagens de arquivos `.pptx`
- Identificação e classificação de legendas associadas a cada imagem
- Extração inteligente de metadados (data, pavimento, tipo de serviço, categoria técnica)
- Conversão das imagens para base64, ajustadas em altura padrão (234px) para visualização no Power BI
- Exportação para planilha `.xlsx` com base pronta para uso

## Tecnologias utilizadas

- Python 3.x
- pandas
- Pillow (manipulação de imagens)
- python-pptx (leitura de slides)
- openpyxl (salvar Excel)

## Estrutura do projeto

```
bi-photo-automation/
├── README.md
├── LICENSE
├── requirements.txt
│
├── src/
│   └── main.py
│
├── examples/
│   ├── input_images/         # imagens genéricas de exemplo
│   └── output_dataset.csv    # base gerada simulada para visualização
```

## Como usar

1. Coloque arquivos `.pptx` com imagens e legendas em uma pasta de entrada.
2. Execute `main.py` após configurar os caminhos no início do script.
3. A planilha de saída conterá os dados processados, com uma coluna de imagem em base64.
4. Conecte o arquivo `.xlsx` no Power BI e use um visualizador de imagem customizado para exibir os registros.

## Observação

- Este repositório utiliza dados simulados e imagens fictícias para preservar a confidencialidade de projetos reais.
- Certifique-se de anonimizar quaisquer imagens ou nomes antes de uso público.

---

Projeto com foco em eficiência operacional, padronização e integração entre áreas técnicas, planejamento e BI.
