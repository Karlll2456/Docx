# Gerador de Documentos DOCX

Este projeto contém um script Python que cria documentos do Microsoft Word (.docx) de forma programática.

## 📋 Pré-requisitos

- Python 3.6 ou superior
- pip (gerenciador de pacotes do Python)

## 🚀 Como executar

### 1. Clone o repositório

```bash
git clone https://github.com/Karlll2456/Docx.git
cd Docx
```

### 2. Instale as dependências

```bash
pip install -r requirements.txt
```

Ou instale diretamente:

```bash
pip install python-docx
```

### 3. Execute o script

```bash
python create_document.py
```

O arquivo `exemplo.docx` será criado **no mesmo diretório** onde você executou o script.

## 📄 O que o script faz

O script `create_document.py` cria um documento Word de exemplo chamado `exemplo.docx` **no diretório atual** que inclui:

- ✅ Título centralizado
- ✅ Parágrafos com texto formatado (negrito, itálico, cores)
- ✅ Listas com marcadores
- ✅ Tabelas com dados
- ✅ Múltiplas seções com subtítulos

## 📦 Estrutura do Projeto

```
Docx/
├── create_document.py   # Script principal para criar documentos
├── requirements.txt      # Dependências do projeto
├── README.md            # Este arquivo
└── exemplo.docx         # Documento gerado (após execução, no diretório local)
```

**Nota:** O arquivo `exemplo.docx` não aparece no repositório Git pois está no `.gitignore`. Ele é criado localmente quando você executa o script.

## 🛠️ Personalização

Você pode modificar o arquivo `create_document.py` para:

- Alterar o conteúdo do documento
- Adicionar mais formatação
- Incluir imagens
- Criar diferentes estilos
- Gerar múltiplos documentos

## 📚 Documentação da biblioteca

Para mais informações sobre a biblioteca `python-docx`, consulte:
- [Documentação oficial](https://python-docx.readthedocs.io/)

## 👤 Autor

Karlll2456

## 📝 Licença

Este projeto é de código aberto e está disponível para uso livre.