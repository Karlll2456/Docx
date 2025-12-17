# Gerador de Documentos DOCX

Este projeto contém scripts Python que criam documentos do Microsoft Word (.docx) de forma programática.

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

### 3. Execute o script desejado

**Para gerar um parecer técnico:**
```bash
python create_document.py
```

**Para gerar a pesquisa sobre IBGE e APM:**
```bash
python create_ibge_research.py
```

## 📄 Scripts Disponíveis

### 1. create_document.py - Parecer Técnico

Este script cria um documento Word chamado `parecer_tecnico.docx` que inclui:

- ✅ Título: "PARECER TÉCNICO - CRIMES NAS CONTESTAÇÕES"
- ✅ Cabeçalho com data
- ✅ Seções: EMENTA, RELATÓRIO, FUNDAMENTAÇÃO, CONCLUSÃO
- ✅ Formatação: Arial 12, justificado, espaçamento 1.5
- ✅ Assinatura

**Uso:**
```bash
# Com argumentos
python create_document.py --ementa "..." --relatorio "..." --fundamentacao "..." --conclusao "..."

# Ou via STDIN
python create_document.py --stdin
```

### 2. create_ibge_research.py - Pesquisa sobre IBGE e APM

Este script cria um documento Word completo chamado `pesquisa_ibge_apm.docx` com informações sobre:

- ✅ O que é o IBGE (Instituto Brasileiro de Geografia e Estatística)
- ✅ História e missão do IBGE
- ✅ Funções e atribuições do instituto
- ✅ Principais censos e pesquisas
- ✅ O cargo de Agente de Pesquisas e Mapeamento (APM)
- ✅ Conteúdo programático completo do concurso para APM
- ✅ Todas as disciplinas: Português, Matemática, Raciocínio Lógico, Ética, Informática e Geografia
- ✅ Dicas de preparação para o concurso
- ✅ Informações sobre remuneração e benefícios

**Uso:**
```bash
python create_ibge_research.py
```

## 📦 Estrutura do Projeto

```
Docx/
├── create_document.py        # Script para criar pareceres técnicos
├── create_ibge_research.py   # Script para gerar pesquisa sobre IBGE e APM
├── requirements.txt           # Dependências do projeto
├── README.md                  # Este arquivo
├── parecer_tecnico.docx       # Parecer gerado (após execução)
└── pesquisa_ibge_apm.docx     # Pesquisa gerada (após execução)
```

## 🛠️ Personalização

Você pode modificar os scripts Python para:

- Alterar o conteúdo dos documentos
- Adicionar mais formatação
- Incluir imagens
- Criar diferentes estilos
- Gerar múltiplos documentos
- Adaptar para outros tipos de documentos

### Exemplos de Uso

**Parecer Técnico com seções específicas:**
```bash
python create_document.py \
  --ementa "Texto da ementa aqui" \
  --relatorio "Descrição do relatório" \
  --fundamentacao "Fundamentação legal" \
  --conclusao "Conclusão do parecer"
```

**Pesquisa IBGE/APM:**
```bash
python create_ibge_research.py
# Gera automaticamente um documento completo com toda a pesquisa
```

## 📚 Documentação da biblioteca

Para mais informações sobre a biblioteca `python-docx`, consulte:
- [Documentação oficial](https://python-docx.readthedocs.io/)

## 👤 Autor

Karlll2456

## 📝 Licença

Este projeto é de código aberto e está disponível para uso livre.