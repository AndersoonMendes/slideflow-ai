# Arquitetura do SlideFlow AI

## 1. Visão geral

A V1 utiliza uma arquitetura sequencial: cada etapa recebe os dados da etapa anterior e produz a entrada necessária para a próxima.

## 2. Fluxo de processamento

```text
Tema + documentos
       |
       v
Extração de texto
       |
       v
Preparação das fontes
       |
       v
Ollama / modelo de linguagem
       |
       v
Roteiro estruturado em JSON
       |
       v
Geração do PowerPoint
       |
       v
Arquivo .pptx
```

## 3. Componentes

### Entrada

Recebe o tema da apresentação e os caminhos dos arquivos de referência.

### Extração de fontes

Converte documentos PDF e TXT em texto processável.

### Geração de conteúdo

Envia o tema e o conteúdo extraído ao modelo de linguagem e solicita um roteiro estruturado em JSON.

### Geração de slides

Converte o roteiro em uma apresentação PowerPoint editável.

### Saída

Armazena o arquivo final no diretório `saida/`.

## 4. Decisões técnicas

- **Python:** linguagem principal para integração com modelos de linguagem e geração de documentos.
- **Ollama:** execução local do modelo de linguagem.
- **JSON:** formato intermediário entre a geração de conteúdo e a criação dos slides.
- **python-pptx:** criação de apresentações editáveis.
- **Arquitetura sequencial:** reduz a complexidade inicial e facilita o diagnóstico de erros.

## 5. Limitações arquiteturais

Na V1, o fluxo não possui agentes autônomos, memória persistente de execução, mecanismos de revisão independentes ou recuperação automática de falhas.

## 6. Evolução planejada

A arquitetura poderá evoluir para agentes especializados em pesquisa documental, elaboração do roteiro, geração de slides e revisão.

Um orquestrador será responsável por coordenar esses agentes, validar resultados e controlar o fluxo de execução.
