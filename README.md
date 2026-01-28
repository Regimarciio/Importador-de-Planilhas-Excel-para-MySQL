# 📊 Importador de Planilhas Excel para MySQL

Este projeto é um **script Python de ETL (Extração, Transformação e Carga)** que lê uma planilha Excel, higieniza os dados, gera um identificador único por OS e importa as informações automaticamente para um banco de dados **MySQL**, realizando **INSERT ou UPDATE** conforme necessário.

O script foi desenvolvido para cenários reais de controle de **ordens de serviço (OS)**, acompanhamento técnico e integração de dados administrativos.

---

## 🚀 Funcionalidades

* 📂 Leitura automática de arquivos **Excel (.xlsx)**
* 🧹 Limpeza e padronização dos nomes das colunas
* 🗑 Remoção automática de colunas `Unnamed`
* 🔢 Conversão segura da coluna **OS** para inteiro
* 🆔 Geração de **número único** baseado em:

  * Ano corrente (2 dígitos)
  * Número da OS
  * Sufixo incremental em caso de duplicidade
* 🏗 Criação automática da tabela no MySQL
* 🔁 Inserção ou atualização automática com `ON DUPLICATE KEY UPDATE`
* 📊 Relatório final com total de registros inseridos, atualizados e ignorados
* 🔒 Fechamento seguro da conexão com o banco

---

## 🧠 Lógica do Número Único

O campo `numero_unico` é gerado no formato:

```
AAOOOO
```

Onde:

* `AA` → Ano atual (ex: 25)
* `OOOO` → Número da OS com 4 dígitos
* Caso haja duplicidade, é adicionado um sufixo incremental (`2501231`, `2501232`, etc.)

Este campo é definido como **UNIQUE** no banco de dados.

---

## 🛠 Tecnologias Utilizadas

* **Python 3.10+**
* **Pandas**
* **MySQL Connector**
* **MySQL / MariaDB**

---

## 📦 Requisitos

Instale as dependências com:

```bash
pip install pandas mysql-connector-python openpyxl
```

---

## ⚙ Configuração

Edite no script as configurações abaixo conforme seu ambiente:

```python
arquivo_excel = r"CAMINHO_DO_ARQUIVO.xlsx"

conn = mysql.connector.connect(
    host="localhost",
    user="root",
    password="",
    database="oficina_db",
    charset="utf8mb4"
)
```

> ⚠ Recomenda-se o uso de variáveis de ambiente (`.env`) em ambientes de produção.

---

## ▶ Como Executar

```bash
python LeitordePlanilhas.py
```

Ao final da execução, será exibido um resumo como:

```
📊 Importação concluída!
- Inseridos: 120
- Atualizados: 15
- Ignorados: 2
```

---

## 🧩 Estrutura da Tabela

A tabela é criada automaticamente com:

* `id` → chave primária auto incremento
* `os` → número da ordem de serviço
* `numero_unico` → identificador único
* Datas convertidas para `DATE`
* Valores monetários em `DECIMAL(15,2)`
* Demais campos em `VARCHAR(255)`

---

## 📈 Possíveis Melhorias Futuras

* 🔍 Validação automática contra dados já existentes no banco
* 📄 Suporte a arquivos CSV
* 🖥 Interface gráfica (Tkinter)
* 📝 Logs em arquivo
* 🔐 Configuração via `.env`
* ⏱ Execução automatizada (Task Scheduler / Cron)

---

## 👨‍💻 Autor

**Reginaldo Marcilio**
Projeto desenvolvido para automação de processos administrativos e técnicos.

---

## 📄 Licença

Este projeto é de uso livre para fins educacionais e profissionais.
Sinta-se à vontade para adaptar e evoluir conforme sua necessidade.

---

⭐ Se este projeto te ajudou, considere deixar uma estrela no repositório!
