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
from datetime import datetime

app = Flask(__name__)
CORS(app)  # Permite requisições de qualquer origem

# Arquivo de usuários (pode ser substituído por banco de dados)
USERS_FILE = "users_server.json"

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
    """Salva usuários no arquivo."""
    try:
        with open(USERS_FILE, "w", encoding="utf-8") as f:
            json.dump(users, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False

def ensure_default_admin():
    """Cria usuário admin padrão se não existir."""
    users = load_users()
    if "admin" not in users:
        # Hash da senha "admin123"
        password_hash = bcrypt.hashpw("admin123".encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
        users["admin"] = {
            "password_hash": password_hash,
            "is_admin": True,
            "created_at": datetime.now().isoformat()
        }
        save_users(users)
        print("Usuário admin padrão criado: admin/admin123")

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
    """Retorna lista de todos os usuários (requer autenticação de admin)."""
    # Em produção, adicionar autenticação do admin aqui
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
    """Adiciona um novo usuário."""
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
    """Atualiza um usuário existente."""
    data = request.json
    password = data.get('password', '').strip()
    is_admin = data.get('is_admin', None)
    
    users = load_users()
    if username not in users:
        return jsonify({"success": False, "message": "Usuário não encontrado"}), 404
    
    if username == "admin" and is_admin is False:
        return jsonify({"success": False, "message": "Não é possível remover privilégios de admin"}), 403
    
    if password:
        users[username]["password_hash"] = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    
    if is_admin is not None:
        users[username]["is_admin"] = is_admin
    
    if save_users(users):
        return jsonify({"success": True, "message": "Usuário atualizado com sucesso"})
    
    return jsonify({"success": False, "message": "Erro ao salvar usuário"}), 500

@app.route('/api/users/<username>', methods=['DELETE'])
def delete_user(username):
    """Remove um usuário."""
    if username == "admin":
        return jsonify({"success": False, "message": "Não é possível remover o usuário admin"}), 403
    
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
    print("Usuário admin padrão: admin / admin123")
    print("=" * 60)
    
    # Roda o servidor em todas as interfaces (acessível na rede local)
    app.run(host='0.0.0.0', port=5000, debug=True)
