
from flask import Flask, render_template, request, jsonify, session, redirect, url_for, flash
from consultar_planilha import Planilha
import random
import os
import logging

app = Flask(__name__)
app.secret_key = os.urandom(24)

# --- TESTE SSSSS   CONFIGURAÇÃO DE LOG DE ATIVIDADES ---
log_formatter = logging.Formatter('%(asctime)s | %(message)s', datefmt='%d/%m/%Y %H:%M:%S')
log_handler = logging.FileHandler('activity.log', encoding='utf-8')
log_handler.setFormatter(log_formatter)
activity_logger = logging.getLogger('activity_logger')
activity_logger.setLevel(logging.INFO)
if not activity_logger.handlers:
    activity_logger.addHandler(log_handler)
# --- FIM DA CONFIGURAÇÃO DE LOG ---

planilha_leitor = Planilha()

# --- ROTAS DE AUTENTICAÇÃO ---
@app.route('/', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form.get('username')
        password = request.form.get('password')
        captcha_answer = request.form.get('captcha')

        try:
            captcha_int = int(captcha_answer)
        except (ValueError, TypeError):
            flash('Resposta da verificação inválida.', 'error')
            return redirect(url_for('login'))

        if captcha_int != session.get('captcha_solution'):
            flash('Resposta da verificação incorreta.', 'error')
            return redirect(url_for('login'))

        usuarios_permitidos = planilha_leitor.carregar_usuarios()
        user_data = usuarios_permitidos.get(str(username))

        if user_data and str(user_data.get('password')) == password:
            session['logged_in'] = True
            session['username'] = user_data['username']
            session['user_fullname'] = user_data.get('name', 'Utilizador')
            activity_logger.info(f"LOGIN_SUCCESS | {session['user_fullname']} ({username}) | Acesso concedido")
            flash(f"Bem-vindo, {session['user_fullname']}!", 'info')
            return redirect(url_for('consulta'))
        else:
            if not user_data:
                activity_logger.warning(f"LOGIN_FAIL | {username} | Utilizador não encontrado")
            else:
                activity_logger.warning(f"LOGIN_FAIL | {username} | Senha incorreta")
            flash('Utilizador ou senha inválidos.', 'error')
            return redirect(url_for('login'))

    num1 = random.randint(1, 10)
    num2 = random.randint(1, 10)
    session['captcha_solution'] = num1 + num2
    captcha_question = f"Quanto é {num1} + {num2}?"
    return render_template('login.html', captcha_question=captcha_question)

@app.route('/logout')
def logout():
    user_fullname = session.get('user_fullname', 'Desconhecido')
    username = session.get('username', 'desconhecido')
    activity_logger.info(f"LOGOUT | {user_fullname} ({username}) | Sessão terminada")
    session.clear()
    flash('Sessão terminada com sucesso.', 'info')
    return redirect(url_for('login'))

# --- ROTA DA APLICAÇÃO PRINCIPAL (Protegida) ---
@app.route('/consulta')
def consulta():
    if not session.get('logged_in'):
        flash('Por favor, faça o login para aceder a esta página.', 'error')
        return redirect(url_for('login'))
    return render_template('index.html')

# --- ROTAS DA API (Protegidas) ---
@app.route('/buscar', methods=['POST'])
def buscar():
    if not session.get('logged_in'):
        return jsonify({'status': 'erro', 'mensagem': 'Acesso não autorizado'}), 401
    
    termo = request.json.get('termo')
    user_fullname = session.get('user_fullname', 'Desconhecido')
    username = session.get('username', 'desconhecido')
    activity_logger.info(f"SEARCH | {user_fullname} ({username}) | Termo: '{termo}'")
    
    item = planilha_leitor.buscar_item(termo)

    if item:
        # Normaliza os campos numéricos de forma segura
        for campo in ['SALDO ESTOQUE', 'MEDIA MENSAL', 'EM COMPRA', 'PROJECAO DO ESTOQUE (DIAS)']:
            campo_real = next((k for k in item.keys() if k.upper().strip() == campo), None)
            if campo_real:
                item[campo_real] = planilha_leitor._coerce_number(item.get(campo_real, '0'))
        return jsonify({'status': 'sucesso', 'data': item})

    return jsonify({'status': 'erro', 'mensagem': 'Código ou descrição não encontrados.'})

# --- ROTA DE SUGESTÕES ATUALIZADA ---
@app.route('/sugestoes')
def sugestoes():
    if not session.get('logged_in'):
        return jsonify([]), 401
    try:
        sugestoes_lista = planilha_leitor.obter_sugestoes()
        if not sugestoes_lista:
             # NOVO: Log para indicar que a lista de sugestões veio vazia
             activity_logger.warning("SUGGESTION_WARN | A lista de sugestões retornou vazia. Verifique os nomes das colunas na planilha.")
        return jsonify(sugestoes_lista)
    except Exception as e:
        activity_logger.error(f"SUGGESTION_ERROR | Erro ao buscar sugestões: {e}")
        return jsonify([])

@app.route('/suprimentos', methods=['POST'])
def get_suprimentos():
    if not session.get('logged_in'):
        return jsonify({'solicitacoes': [], 'compras': []}), 401
        
    codigo = request.json.get('codigo')
    if codigo:
        dados = planilha_leitor.buscar_suprimentos(codigo)
        return jsonify(dados)
    return jsonify({'solicitacoes': [], 'compras': []})

if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=8080)

