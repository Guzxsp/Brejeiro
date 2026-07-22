# -*- coding: utf-8 -*-
"""
Servidor de Autenticação Centralizado
API Flask para gerenciamento de usuários do sistema Certidões ONR
"""

from flask import Flask, request, jsonify
from flask_cors import CORS
import bcrypt
import json
import os
import shutil
from datetime import datetime
import secrets

app = Flask(__name__)
CORS(app)  # Permite requisições de qualquer origem

# Token secreto para proteger rotas admin (deve ser configurado via variável de ambiente)
ADMIN_SECRET_TOKEN = os.environ.get('ADMIN_SECRET_TOKEN', 'brejeiro_admin_secret_2024_secure_token_change_in_production')

# Arquivo de usuários (pode ser substituído por banco de dados)
USERS_FILE = "users_server.json"
BACKUP_DIR = "backups_users"

# Configuração de backup
MAX_BACKUPS = 10  # Mantém apenas os 10 backups mais recentes

def load_users():
    """Carrega usuários do arquivo."""
    try:
        if os.path.exists(USERS_FILE):
            with open(USERS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return {}

def save_users(users):
    """Salva usuários no arquivo e faz backup automático."""
    try:
        # Faz backup antes de salvar
        backup_users()
        
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(users, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False

def backup_users():
    """Faz backup do arquivo de usuários."""
    try:
        if not os.path.exists(USERS_FILE):
            return
        
        # Cria diretório de backup se não existir
        if not os.path.exists(BACKUP_DIR):
            os.makedirs(BACKUP_DIR)
        
        # Nome do arquivo de backup com timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = os.path.join(BACKUP_DIR, f"users_backup_{timestamp}.json")
        
        # Copia arquivo atual para backup
        shutil.copy2(USERS_FILE, backup_file)
        
        # Remove backups antigos, mantendo apenas os mais recentes
        backups = sorted([f for f in os.listdir(BACKUP_DIR) if f.startswith("users_backup_")])
        while len(backups) > MAX_BACKUPS:
            old_backup = os.path.join(BACKUP_DIR, backups.pop(0))
            os.remove(old_backup)
            
        print(f"Backup criado: {backup_file}")
    except Exception as e:
        print(f"Erro ao criar backup: {e}")

def verify_admin_token():
    """Verifica se o token de admin está presente no header."""
    token = request.headers.get('X-Admin-Token')
    if not token:
        return False
    return secrets.compare_digest(token, ADMIN_SECRET_TOKEN)

def ensure_default_admin():
    """Cria usuário admin padrão se não existir."""
    users = load_users()
    if "Comprasoja" not in users:
        # Hash da senha "compras001"
        password_hash = bcrypt.hashpw("compras001".encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
        users["Comprasoja"] = {
            "password_hash": password_hash,
            "is_admin": True,
            "created_at": datetime.now().isoformat()
        }
        save_users(users)
        print("Usuário admin padrão criado: Comprasoja/compras001")

# Inicializa admin ao iniciar
ensure_default_admin()

@app.route('/health', methods=['GET'])
def health_check():
    """Endpoint para verificar se o servidor está online."""
    return jsonify({"status": "online", "timestamp": datetime.now().isoformat()})

@app.route('/api/auth/login', methods=['POST'])
def login():
    """Autentica um usuário."""
    data = request.json
    username = data.get('username', '').strip()
    password = data.get('password', '').strip()
    
    if not username or not password:
        return jsonify({"success": False, "message": "Usuário e senha são obrigatórios"}), 400
    
    users = load_users()
    if username not in users:
        return jsonify({"success": False, "message": "Usuário ou senha incorretos"}), 401
    
    user_data = users[username]
    stored_hash = user_data["password_hash"].encode('utf-8')
    
    # Verifica a senha
    if bcrypt.checkpw(password.encode('utf-8'), stored_hash):
        # Retorna dados do usuário sem a senha
        return jsonify({
            "success": True,
            "user": {
                "username": username,
                "is_admin": user_data.get("is_admin", False)
            }
        })
    
    return jsonify({"success": False, "message": "Usuário ou senha incorretos"}), 401

@app.route('/api/users', methods=['GET'])
def get_users():
    """Retorna lista de todos os usuários (requer token de admin)."""
    if not verify_admin_token():
        return jsonify({"success": False, "message": "Token de admin inválido ou ausente"}), 403
    
    users = load_users()
    user_list = []
    for username, data in users.items():
        user_list.append({
            "username": username,
            "is_admin": data.get("is_admin", False),
            "created_at": data.get("created_at", "")
        })
    return jsonify({"success": True, "users": user_list})

@app.route('/api/users', methods=['POST'])
def add_user():
    """Adiciona um novo usuário (requer token de admin)."""
    if not verify_admin_token():
        return jsonify({"success": False, "message": "Token de admin inválido ou ausente"}), 403
    
    data = request.json
    username = data.get('username', '').strip()
    password = data.get('password', '').strip()
    is_admin = data.get('is_admin', False)
    
    if not username or not password:
        return jsonify({"success": False, "message": "Usuário e senha são obrigatórios"}), 400
    
    users = load_users()
    if username in users:
        return jsonify({"success": False, "message": "Usuário já existe"}), 409
    
    # Hash da senha
    password_hash = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    
    users[username] = {
        "password_hash": password_hash,
        "is_admin": is_admin,
        "created_at": datetime.now().isoformat()
    }
    
    if save_users(users):
        return jsonify({"success": True, "message": "Usuário criado com sucesso"})
    
    return jsonify({"success": False, "message": "Erro ao salvar usuário"}), 500

@app.route('/api/users/<username>', methods=['PUT'])
def update_user(username):
    """Atualiza um usuário existente (requer token de admin)."""
    if not verify_admin_token():
        return jsonify({"success": False, "message": "Token de admin inválido ou ausente"}), 403
    
    data = request.json
    password = data.get('password', '').strip()
    is_admin = data.get('is_admin', None)
    
    users = load_users()
    if username not in users:
        return jsonify({"success": False, "message": "Usuário não encontrado"}), 404
    
    if username == "Comprasoja" and is_admin is False:
        return jsonify({"success": False, "message": "Não é possível remover privilégios de admin do usuário principal"}), 403
    
    if password:
        users[username]["password_hash"] = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    
    if is_admin is not None:
        users[username]["is_admin"] = is_admin
    
    if save_users(users):
        return jsonify({"success": True, "message": "Usuário atualizado com sucesso"})
    
    return jsonify({"success": False, "message": "Erro ao salvar usuário"}), 500

@app.route('/api/users/<username>', methods=['DELETE'])
def delete_user(username):
    """Remove um usuário (requer token de admin)."""
    if not verify_admin_token():
        return jsonify({"success": False, "message": "Token de admin inválido ou ausente"}), 403
    
    if username == "Comprasoja":
        return jsonify({"success": False, "message": "Não é possível remover o usuário admin principal"}), 403
    
    users = load_users()
    if username not in users:
        return jsonify({"success": False, "message": "Usuário não encontrado"}), 404
    
    del users[username]
    
    if save_users(users):
        return jsonify({"success": True, "message": "Usuário removido com sucesso"})
    
    return jsonify({"success": False, "message": "Erro ao remover usuário"}), 500

if __name__ == '__main__':
    print("=" * 60)
    print("Servidor de Autenticação - Certidões ONR")
    print("=" * 60)
    print("Servidor iniciado em: http://localhost:5000")
    print("Para acessar de outros computadores, use o IP da máquina")
    print("Usuário admin padrão: Comprasoja / compras001")
    print(f"Token de admin: {ADMIN_SECRET_TOKEN}")
    print("=" * 60)
    
    # Roda o servidor em todas as interfaces (acessível na rede local)
    app.run(host='0.0.0.0', port=5000, debug=True)
