# FiveCom ERP Enterprise V1

ERP white-label de estoque e patrimônio, em português, com responsividade mobile.

## Acesso inicial
- Usuário: `admin`
- Senha: `Admin@123`

Altere a senha assim que entrar.

## Execução local
```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\Scripts\activate
pip install -r requirements.txt
python app.py
```
Abra http://localhost:8000

## Docker
```bash
docker compose up -d --build
```
Abra http://localhost:8000

## Render
Envie este projeto ao GitHub e crie um Blueprint usando `render.yaml`.

## Recursos
- Dashboard, tarefas e notificações
- Produtos, perecíveis, lotes, validade e FEFO
- Entradas/saídas e estoque automático
- Fornecedores, material e patrimônio, base de conhecimento
- Assistentes de automações, gráficos e relatórios
- PDF, Excel e Word
- White-label: nome, aba, logos, favicon, cores e posição do login
- Usuários, perfis, permissões, ocultação de módulos e auditoria
- Anexos com URL automática
- Menu responsivo

## Observação
Para produção com múltiplas instâncias, utilize PostgreSQL e armazenamento de objetos para uploads.

## Refinamento V1.1
Inclui categorias, filtro de produtos, fotos, leitura de código de barras pela câmera em dispositivos compatíveis, contas detalhadas de fornecedores, gráficos funcionais, gráficos em relatórios e dados fictícios completos.
