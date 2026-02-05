<p align="center"> <img width="300" height="230" alt="Catalogo+" src="https://github.com/user-attachments/assets/38f01a77-f234-4406-b680-40a9acb494b7" /> </p>

---

## *Catálogo+* é uma solução full-stack desenvolvida para automatizar a criação de catálogos de moda e têxtil. O sistema integra uma interface web moderna com o motor de renderização do Adobe InDesign, permitindo que usuários gerem PDFs prontos para impressão ou distribuição digital em questão de minutos, eliminando o trabalho manual de diagramação.

---
# 🎞️ Vídeos

> ## 💻 Desktop
https://github.com/user-attachments/assets/f2bee4d8-28fb-46c7-9bc5-d925c1faf680

> ## 📱 Mobile

https://github.com/user-attachments/assets/cb3d139c-279a-461f-b13f-443734bdb999


---
<br>
<br>



## 🎯 Objetivo do Projeto
 
 - ### O objetivo principal é otimizar o fluxo de trabalho do setor de criação, transformando dados brutos (referências, preços, imagens) em layouts complexos automaticamente. O sistema resolve o problema de gargalo na produção de catálogos, permitindo geração sob demanda através de uma fila de processamento segura.
<br>


## 🛠️ Arquitetura e Tecnologias

### O projeto opera em um ambiente híbrido, funcionando como uma aplicação Web local.

> ## Backend
- Python 3 & Flask: Núcleo da aplicação web e gerenciamento de rotas.

- Threading & Locks: Implementação de Fila de Espera (Queue) e Bloqueio de Sessão (check_session_queue) para gerenciar o acesso único ao InDesign, prevenindo conflitos de concorrência.

- Subprocess & VBScript: Ponte de comunicação entre o Python e o Windows Script Host para invocar o InDesign.

- Pandas: Manipulação e geração de arquivos CSV para o Data Merge.

> ## Automação (Scripting)
- Adobe InDesign Scripting (JSX/ExtendScript): Scripts dedicados para manipular o DOM do InDesign, realizar a mesclagem de dados (Data Merge), exportar PDFs e limpar a memória.

- Templates (.indd): Arquivos mestres configurados com placeholders dinâmicos.

> ## Frontend
- HTML5 / CSS3: Interface responsiva para seleção de layouts e input de dados.

- JavaScript: Lógica de client-side.

- SweetAlert2: Sistema de notificações e alertas amigáveis para feedback de erros e status da fila.

# 🚀 Funcionalidades Chave
-  Geração Modular: Capacidade de gerar apenas Capa, Miolo (Produtos), Contra-capa ou o Catálogo Completo.

-  Fila de Processamento Inteligente:

-  Sistema de Lock global: Impede que dois usuários usem o InDesign simultaneamente.

-  Sala de Espera: Usuários secundários aguardam em uma tela de loading até que o motor de renderização esteja livre.

-  Protocolo de Recuperação (Kill Switch): Monitoramento de Timeouts. Se o InDesign travar (ex: exceder 5 minutos), o sistema encerra forçadamente os processos (taskkill) e libera a fila automaticamente.

-  Timeout de Sessão: Redirecionamento automático de usuários inativos para liberar recursos.

# ⚙️ Pré-requisitos
## Para rodar este projeto, o ambiente deve atender aos seguintes requisitos estritos:

- Sistema Operacional: Windows 10 ou 11 (Obrigatório devido ao uso de VBS/COM Objects).

- Software: Adobe InDesign (Versão 2024 ou superior recomendada) instalado e licenciado.

- Linguagem: Python 3.10+.

- Dependências: Listadas em requirements.txt.

# 📦 Instalação e Configuração
## 1.  Clone o repositório:

```bash
git clone https://github.com/seu-usuario/catalogo-plus.git
cd catalogo-plus
```

## 2.  Crie e ative um ambiente virtual:

```bash
python -m venv venv
# No Windows:
venv\Scripts\activate
```

## 3. Instale as dependências:

```bash
pip install -r requirements.txt
```

## 4. Configuração do Banco de Dados:

- Certifique-se de que a conexão MySQL está configurada corretamente no app.py (ou variáveis de ambiente).

- Altere os dados do BD para os correspondentes ao seus.
  
## 5. Verifique os Caminhos:

- Confira se os caminhos absolutos para os scripts JSX e Templates dentro do app.py correspondem à estrutura da sua máquina.

# ▶️ Como Usar
## 1. Inicie o servidor Flask:

```bash
python app.py
# Acesse no navegador: http://localhost:5000
```


## 📂 Estrutura de Pastas (Resumo)
```
catalogo_plus/
├── app.py                 # Aplicação principal (Flask)
├── requirements.txt       # Dependências Python
├── static/                # Assets (Imagens, CSS, JS, Fontes)
├── templates/             # Arquivos HTML (Jinja2)
│   ├── index.html
│   ├── painel.html
│   ├── visualizer.html
│   └── waiting.html       # Tela de fila de espera
└── indesign/              # Núcleo da Automação
    ├── CSV/               # Arquivos de dados gerados
    ├── output/            # PDFs finais gerados
    ├── template/          # Arquivos .indd base
    ├── script_capa.jsx    # Script de automação da Capa
    ├── script_produto.jsx # Script de automação do Miolo
    └── ...
```

## ⚠️ Notas Importantes
- Single-Threaded por Design: O Adobe InDesign não é um serviço de servidor multi-thread. O sistema foi desenhado para enfileirar requisições. Não tente rodar múltiplas instâncias do InDesign manualmente.

- Fontes: As fontes utilizadas nos templates (ex: Parisienne, Agenda) devem estar instaladas no Windows para que o InDesign as reconheça.

---

---
<p align="center">Desenvolvido por Yasmin Mamud</p>
<br>
<div align="center">
  <a href="mailto:yasmin.mamud299@gmail.com"><img src="https://img.shields.io/badge/Gmail-D14836?style=for-the-badge&logo=gmail&logoColor=white" target="_blank"></a>
  <a href="https://www.linkedin.com/in/yasmin-mamud299" target="_blank"><img src="https://img.shields.io/badge/-LinkedIn-%230077B5?style=for-the-badge&logo=linkedin&logoColor=white" target="_blank"></a>
  <a href="https://www.instagram.com/euymyh" target="_blank"><img src="https://img.shields.io/badge/-Instagram-%23E4405F?style=for-the-badge&logo=instagram&logoColor=white" target="_blank"></a>
</div>


  
