# SlideFlow AI

Sistema de geração automática de apresentações PowerPoint a partir de um tema e de fontes documentais, utilizando um modelo de linguagem executado localmente.

## Objetivo

O SlideFlow AI tem como objetivo transformar materiais de referência em apresentações organizadas, editáveis e fáceis de revisar.

O projeto começa com uma implementação simples e evolui gradualmente para uma arquitetura de orquestração de agentes especializados.

## Funcionalidades — V1

- Extração de texto de arquivos PDF e TXT.
- Geração de um roteiro de apresentação com um modelo de linguagem local.
- Criação de arquivos PowerPoint editáveis (`.pptx`).
- Inclusão de referências às fontes fornecidas.
- Execução local utilizando Ollama.

## Tecnologias

- Python
- Ollama
- Qwen 3.5 4B
- PyMuPDF
- python-pptx
- Requests

## Arquitetura inicial

1. **Entrada:** recebe o tema e os caminhos dos documentos.
2. **Extração:** converte os documentos em texto.
3. **Geração:** solicita ao modelo de linguagem um roteiro estruturado.
4. **Apresentação:** transforma o roteiro em slides editáveis.
5. **Saída:** salva o PowerPoint gerado.

## Requisitos

- Python 3.10 ou superior, sujeito à compatibilidade das dependências.
- Ollama instalado e em execução.
- Modelo `qwen3.5:4b` disponível no Ollama.

## Instalação

Clone o repositório e entre na pasta do projeto.

Crie e ative um ambiente virtual:

```bash
python3 -m venv .venv
source .venv/bin/activate
```

Instale as dependências:

```bash
pip install -r requirements.txt
```

Verifique se o Ollama está acessível e se o modelo está disponível:

```bash
ollama list
```

## Uso

Execute o programa informando o tema e uma ou mais fontes:

```bash
python main.py "Diferenças entre CPU e GPU" fontes/computacao.txt
```

O arquivo PowerPoint será salvo em `saida/apresentacao.pptx`.

## Limitações conhecidas

- A qualidade depende das fontes e das respostas do modelo.
- Documentos extensos podem ultrapassar o limite de texto processado.
- As referências geradas precisam ser verificadas.
- A V1 não possui pesquisa automática na internet nem agentes independentes.
- PDFs digitalizados podem exigir OCR, ainda não implementado.

## Roadmap

- [x] Definir a proposta da V1.
- [ ] Gerar apresentações a partir de arquivos TXT.
- [ ] Adicionar suporte a PDF.
- [ ] Validar e melhorar o conteúdo gerado.
- [ ] Implementar testes automatizados.
- [ ] Separar o código em módulos.
- [ ] Evoluir para uma arquitetura de agentes.
- [ ] Adicionar interface para usuários.

## Status

Em desenvolvimento — versão inicial.

## Licença

A licença será definida antes da primeira publicação oficial.

## Aviso

O conteúdo gerado por inteligência artificial deve ser revisado antes de ser utilizado em contextos acadêmicos ou profissionais.
