# Sistema de Autenticação Centralizado - Certidões ONR

## Visão Geral

Este sistema permite que múltiplos clientes do aplicativo Certidões ONR se conectem a um servidor centralizado para autenticação e gerenciamento de usuários. Isso garante que:

- **Segurança**: Senhas são criptografadas com bcrypt no servidor
- **Centralização**: Todos os clientes compartilham os mesmos usuários
- **Proteção**: Usuários não podem alterar o arquivo local de usuários
- **Sincronização**: Mudanças feitas em um cliente refletem em todos

## Arquitetura

```
┌─────────────────┐         HTTP/REST API         ┌─────────────────┐
│  Cliente 1      │ ──────────────────────────────> │                 │
│  (Windows PC)   │                                 │  Servidor Flask │
└─────────────────┘                                 │  (auth_server)  │
                                                     │                 │
┌─────────────────┐                                 │  - users_server │
│  Cliente 2      │ ──────────────────────────────> │    .json       │
│  (Outro PC)     │                                 │  - bcrypt       │
└─────────────────┘                                 └─────────────────┘
```

## Configuração do Servidor

### 1. Instalar Dependências

No servidor (pode ser qualquer máquina na rede):

```cmd
cd "C:\Users\gusta\Downloads\donwloads 1\Brejeiroconsultar"
pip install -r requirements_server.txt
```

### 2. Iniciar o Servidor

```cmd
python auth_server.py
```

O servidor vai iniciar em `http://localhost:5000` e será acessível na rede local.

### 3. Configurar Firewall

Certifique-se de que a porta 5000 está liberada no firewall do Windows.

## Configuração dos Clientes

### 1. Primeiro Uso

Ao abrir o aplicativo pela primeira vez, ele vai tentar conectar ao servidor padrão (`http://localhost:5000`).

### 2. Configurar URL do Servidor

Se o servidor estiver em outra máquina:

1. Clique no botão **"🔧 Servidor"** no cabeçalho
2. Insira a URL do servidor, por exemplo:
   - `http://192.168.1.100:5000` (substitua pelo IP real do servidor)
   - `http://localhost:5000` (se servidor e cliente estão na mesma máquina)
3. Clique em **"🧪 Testar"** para verificar a conexão
4. Clique em **"💾 Salvar"** para salvar a configuração

### 3. Login

Use o usuário admin padrão:
- **Usuário**: `admin`
- **Senha**: `admin123`

## Gerenciamento de Usuários

### Acessar o Painel de Admin

1. Faça login com um usuário administrador
2. Clique no botão **"⚙️ Admin"** no cabeçalho

### Criar Novo Usuário

1. No painel de admin, clique em **"➕ Novo Usuário"**
2. Preencha:
   - Usuário
   - Senha
   - Marque "Usuário Administrador" se necessário
3. Clique em **"💾 Salvar"**

### Editar Usuário

1. Selecione um usuário na lista
2. Clique em **"✏️ Editar"**
3. Altere senha ou privilégios de admin
4. Clique em **"💾 Salvar"**

### Excluir Usuário

1. Selecione um usuário na lista
2. Clique em **"🗑️ Excluir"**
3. Confirme a exclusão

**Nota**: O usuário admin não pode ser excluído e não pode perder privilégios de admin.

## Segurança

### Criptografia de Senhas

- Todas as senhas são criptografadas com bcrypt antes de serem salvas
- O servidor nunca armazena senhas em texto plano
- Mesmo com acesso ao arquivo `users_server.json`, não é possível recuperar as senhas originais

### Proteção do Arquivo

- O arquivo `users_server.json` fica apenas no servidor
- Clientes não têm acesso direto a este arquivo
- Todas as operações passam pela API REST

### Recomendações

1. **Mude a senha do admin** imediatamente após o primeiro uso
2. **Use HTTPS em produção** (requer configuração adicional com certificado SSL)
3. **Mantenha o servidor em uma máquina segura** na rede
4. **Faça backup regular** do arquivo `users_server.json`

## Solução de Problemas

### Cliente não consegue conectar

1. Verifique se o servidor está rodando
2. Teste a conexão no navegador: `http://IP_DO_SERVIDOR:5000/health`
3. Verifique se o firewall está bloqueando a porta 5000
4. Use o botão **"🧪 Testar"** no diálogo de configuração

### Indicador mostra "🔴 Offline"

- O cliente não conseguiu conectar ao servidor
- Verifique a URL configurada
- Verifique se o servidor está online

### Erro ao criar/editar usuário

- Verifique a conexão com o servidor
- Verifique se o usuário já existe (para criação)
- Verifique se está tentando remover privilégios do admin

## API Endpoints

O servidor expõe os seguintes endpoints:

- `GET /health` - Verifica se o servidor está online
- `POST /api/auth/login` - Autentica usuário
- `GET /api/users` - Lista todos os usuários
- `POST /api/users` - Cria novo usuário
- `PUT /api/users/<username>` - Atualiza usuário
- `DELETE /api/users/<username>` - Remove usuário

## Arquivos Criados

- `auth_server.py` - Servidor Flask com API de autenticação
- `requirements_server.txt` - Dependências do servidor
- `auth_config.json` - Configuração do cliente (criado automaticamente)
- `users_server.json` - Banco de dados de usuários (no servidor)

## Exemplo de Configuração em Rede

Supondo que o servidor está no IP `192.168.1.100`:

1. **No servidor**:
   ```cmd
   python auth_server.py
   ```

2. **No cliente 1** (IP 192.168.1.50):
   - Configure URL: `http://192.168.1.100:5000`

3. **No cliente 2** (IP 192.168.1.75):
   - Configure URL: `http://192.168.1.100:5000`

Ambos os clientes vão compartilhar os mesmos usuários!
