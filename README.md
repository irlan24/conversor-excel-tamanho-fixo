# 🚀 Conversor Excel para Arquivo de Tamanho Fixo

## 💡 Sobre o Projeto

Desenvolvi uma aplicação desktop em Python que resolve um problema comum em integrações de sistemas: **converter planilhas Excel para arquivos de texto com layout posicional (tamanho fixo)**.

Este formato é amplamente utilizado em:
- 🏦 Integrações bancárias
- 💼 Sistemas de folha de pagamento
- 🔄 Integração com sistemas legados
- 📊 Importação de dados em mainframes

---

## 🤖 Desenvolvimento com Inteligência Artificial

Este projeto foi desenvolvido com o auxílio da **IA Claude (Anthropic)**, que contribuiu significativamente em:

✅ Estruturação da lógica de formatação de campos  
✅ Desenvolvimento da interface gráfica com CustomTkinter  
✅ Implementação das funções de processamento de dados  
✅ Otimização do código e aplicação de boas práticas  
✅ Documentação completa do projeto  

A colaboração com IA permitiu acelerar o desenvolvimento e garantir qualidade no código, demonstrando como a tecnologia pode potencializar a produtividade de desenvolvedores.

---

## ✨ Principais Funcionalidades

🎨 **Interface Gráfica Moderna**
- Design intuitivo com tema escuro
- Desenvolvida com CustomTkinter

⚙️ **Configuração Dinâmica**
- Adicione, remova e reordene colunas em tempo real
- Validação automática das colunas do Excel

🔢 **Dois Tipos de Preenchimento**
- **zfill**: Preenche com zeros à esquerda (CPF, códigos numéricos)
- **ljust**: Preenche com espaços à direita (nomes, descrições)

📊 **Visualização em Tempo Real**
- Preview do tamanho total da linha
- Lista organizada das colunas configuradas

---

## 🖼️ Interface da Aplicação

### Tela Principal
![Tela inicial mostrando seleção de arquivo e campo para adicionar colunas]

### Configuração de Colunas
![Lista de colunas configuradas com opções de ordenação e remoção]

### Resultado da Conversão
![Mensagem de sucesso com informações do arquivo gerado]

---

## 📋 Exemplo Prático

**Entrada (Excel):**
```
CPF          | Nome          | Valor
12345678901  | João Silva    | 1500.50
98765432100  | Maria Santos  | 2300.00
```

**Configuração:**
- CPF: 11 caracteres (zeros à esquerda)
- Nome: 20 caracteres (espaços à direita)
- Valor: 10 caracteres (zeros à esquerda)

**Saída (TXT):**
```
12345678901João Silva          0001500.50
98765432100Maria Santos        0002300.00
```

---

## 🛠️ Tecnologias Utilizadas

- **Python 3.8+** - Linguagem base
- **Pandas** - Manipulação de dados e leitura de Excel
- **CustomTkinter** - Interface gráfica moderna
- **OpenPyXL** - Suporte a arquivos .xlsx

---

## 📦 Como Executar

1. **Instale as dependências:**
```bash
pip install pandas openpyxl customtkinter
```

2. **Execute a aplicação:**
```bash
python conversor_tam_fixo.py
```

3. **Use a interface para:**
   - Selecionar seu arquivo Excel
   - Configurar as colunas desejadas
   - Gerar o arquivo de tamanho fixo

---

## 💭 Reflexão sobre o Uso de IA

O desenvolvimento deste projeto evidenciou como a Inteligência Artificial pode ser uma parceira valiosa na programação:

🎯 **Produtividade**: Redução significativa do tempo de desenvolvimento  
🧠 **Aprendizado**: Exposição a melhores práticas e padrões de código  
🔍 **Qualidade**: Código mais limpo e bem documentado  
⚡ **Agilidade**: Prototipagem rápida de funcionalidades  

A IA não substitui o desenvolvedor, mas potencializa suas capacidades, permitindo foco em aspectos estratégicos e criativos do projeto.

---

## 🔗 Acesse o Código

📂 **GitHub**: [github.com/irlan24/conversor-excel-tamanho-fixo](https://github.com/irlan24/conversor-excel-tamanho-fixo)

⭐ Se você achou útil, deixe uma estrela no repositório!

---

## 📬 Vamos Conversar?

Tem interesse em discutir sobre desenvolvimento com IA, Python ou integração de sistemas?

📧 **Email**: irlan.nonato97@hotmail.com  
💼 **LinkedIn**: linkedin.com/in/irlan24/

---

**#Python #DesenvolvimentoDeSoftware #InteligenciaArtificial #IA #Claude #Automacao #Excel #Programacao #TechInnovation #OpenSource**

---


💻 Desenvolvido com Python | 🤖 Potencializado com IA | ❤️ Compartilhado com a comunidade
