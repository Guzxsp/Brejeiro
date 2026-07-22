# -*- coding: utf-8 -*-
"""
Servidor de Autenticação Centralizado
API Flask para gerenciamento de usuários do sistema Certidões ONR
"""

from flask import Flask, request, jsonify
from flask_cors import CORS
import bcrypt
import os
from datetime import datetime
import secrets
from supabase import create_client

app = Flask(__name__)
CORS(app)  # Permite requisições de qualquer origem

# Cliente Supabase
supabase = create_client(
    os.getenv('SUPABASE_URL'),
    os.getenv('SUPABASE_KEY')
)

# Token secreto para proteger rotas admin (deve ser configurado via variável de ambiente)
ADMIN_SECRET_TOKEN = os.environ.get('ADMIN_SECRET_TOKEN', 'brejeiro_admin_secret_2024_secure_token_change_in_production')

def buscar_usuario(username):
    """Busca um usuário no Supabase pelo username."""
    try:
        resposta = (
            supabase
            .table("users")
            .select("*")
            .eq("username", username)
            .execute()
        )
        
        if resposta.data:
            return resposta.data[0]
        return None
    except Exception as e:
        print(f"Erro ao buscar usuário: {e}")
        return None

def verify_admin_token():
    """Verifica se o token de admin está presente no header."""
    token = request.headers.get('X-Admin-Token')
    if not token:
        return False
    return secrets.compare_digest(token, ADMIN_SECRET_TOKEN)

def ensure_default_admin():
    """Cria usuário admin padrão se não existir."""
    try:
        # Verifica se o usuário já existe
        user = buscar_usuario("Comprasoja")
        if not user:
            # Hash da senha "compras001"
            password_hash = bcrypt.hashpw("compras001".encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
            
            supabase.table("users").insert({
                "username": "Comprasoja",
                "password": password_hash,
                "is_admin": True,
                "created_at": datetime.now().isoformat()
            }).execute()
            print("Usuário admin padrão criado: Comprasoja/compras001")
    except Exception as e:
        print(f"Erro ao criar admin padrão: {e}")

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
    
    user = buscar_usuario(username)
    if not user:
        return jsonify({"success": False, "message": "Usuário ou senha incorretos"}), 401
    
    stored_hash = user["password"].encode('utf-8')
    
    # Verifica a senha
    if bcrypt.checkpw(password.encode('utf-8'), stored_hash):
        # Retorna dados do usuário sem a senha
        return jsonify({
            "success": True,
            "user": {
                "username": username,
                "is_admin": user.get("is_admin", False)
            }
        })
    
    return jsonify({"success": False, "message": "Usuário ou senha incorretos"}), 401

@app.route('/api/users', methods=['GET'])
def get_users():
    """Retorna lista de todos os usuários (requer token de admin)."""
    if not verify_admin_token():
        return jsonify({"success": False, "message": "Token de admin inválido ou ausente"}), 403
    
    try:
        usuarios = (
            supabase
            .table("users")
            .select("username,is_admin,created_at")
            .execute()
        )
        return jsonify({"success": True, "users": usuarios.data})
    except Exception as e:
        return jsonify({"success": False, "message": f"Erro ao buscar usuários: {str(e)}"}), 500

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
    
    # Verifica se usuário já existe
    existing_user = buscar_usuario(username)
    if existing_user:
        return jsonify({"success": False, "message": "Usuário já existe"}), 409
    
    # Hash da senha
    password_hash = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    
    try:
        supabase.table("users").insert({
            "username": username,
            "password": password_hash,
            "is_admin": is_admin,
            "created_at": datetime.now().isoformat()
        }).execute()
        return jsonify({"success": True, "message": "Usuário criado com sucesso"})
    except Exception as e:
        return jsonify({"success": False, "message": f"Erro ao criar usuário: {str(e)}"}), 500

@app.route('/api/users/<username>', methods=['PUT'])
def update_user(username):
    """Atualiza um usuário existente (requer token de admin)."""
    if not verify_admin_token():
        return jsonify({"success": False, "message": "Token de admin inválido ou ausente"}), 403
    
    data = request.json
    password = data.get('password', '').strip()
    is_admin = data.get('is_admin', None)
    
    # Verifica se usuário existe
    user = buscar_usuario(username)
    if not user:
        return jsonify({"success": False, "message": "Usuário não encontrado"}), 404
    
    if username == "Comprasoja" and is_admin is False:
        return jsonify({"success": False, "message": "Não é possível remover privilégios de admin do usuário principal"}), 403
    
    update_data = {}
    if password:
        update_data["password"] = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt()).decode('utf-8')
    if is_admin is not None:
        update_data["is_admin"] = is_admin
    
    try:
        supabase.table("users").update(update_data).eq("username", username).execute()
        return jsonify({"success": True, "message": "Usuário atualizado com sucesso"})
    except Exception as e:
        return jsonify({"success": False, "message": f"Erro ao atualizar usuário: {str(e)}"}), 500

@app.route('/api/users/<username>', methods=['DELETE'])
def delete_user(username):
    """Remove um usuário (requer token de admin)."""
    if not verify_admin_token():
        return jsonify({"success": False, "message": "Token de admin inválido ou ausente"}), 403
    
    if username == "Comprasoja":
        return jsonify({"success": False, "message": "Não é possível remover o usuário admin principal"}), 403
    
    # Verifica se usuário existe
    user = buscar_usuario(username)
    if not user:
        return jsonify({"success": False, "message": "Usuário não encontrado"}), 404
    
    try:
        supabase.table("users").delete().eq("username", username).execute()
        return jsonify({"success": True, "message": "Usuário removido com sucesso"})
    except Exception as e:
        return jsonify({"success": False, "message": f"Erro ao remover usuário: {str(e)}"}), 500

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
