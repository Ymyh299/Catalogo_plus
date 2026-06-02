#codigo app.py
from flask import Flask, render_template, redirect, url_for, session, flash, request, send_file, jsonify
import json
from functools import wraps
from werkzeug.security import check_password_hash
import mysql.connector
import subprocess
import os
import pandas as pd
import time
import tempfile
from concurrent.futures import ThreadPoolExecutor, as_completed
import requests
import mysql.connector
import os
from dotenv import load_dotenv
import fitz
import zipfile
import io
import threading
import uuid as uuid_lib
from werkzeug.security import generate_password_hash
import gspread
from google.oauth2.service_account import Credentials



load_dotenv()

app = Flask(__name__)
app.secret_key = os.getenv('SECRET_KEY')


# CONFIGS
#-------------------------------------------------


DB_CONFIG = { 
    "host": os.getenv('DB_HOST'),
    "user": os.getenv('DB_USER'),
    "password": os.getenv('DB_PASS'),
    "database": os.getenv('DB_NAME')
}

SESSION_LOCK = {
    "user_id": None,  
    "last_active": 0  
}

LOCK_TIMEOUT = 1800

API_URL_BASE = os.getenv('API_URL_BASE')
COMPANY_ID = os.getenv('COMPANY_ID')
APIKEY = os.getenv('API_KEY')


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "indesign")
CSV_DIR = os.path.join(DATA_DIR, "CSV")
os.makedirs(CSV_DIR, exist_ok=True)

creds = Credentials.from_service_account_file(
    os.path.join(BASE_DIR, 'credentials.json'),
    scopes=['https://www.googleapis.com/auth/spreadsheets']
)
gc = gspread.authorize(creds)
sheet = gc.open_by_key('1u6p5gUGHbi1mZxu5qgmwwnpkPieJs9t72a6RN1FUV98')

CSV_PRODUTO_PATH = os.path.join(CSV_DIR, "data_merge_produto.csv")
CSV_CAPA_PATH = os.path.join(CSV_DIR, "data_merge_capa.csv")
CSV_CONTRACAPA_PATH = os.path.join(CSV_DIR, "data_merge_contracapa.csv")


JSX_SCRIPT_COMPLETO = os.path.join(DATA_DIR, "script_completo.jsx")
JSX_SCRIPT_CAPA = os.path.join(DATA_DIR, "script_capa.jsx")
JSX_SCRIPT_CONTRACAPA = os.path.join(DATA_DIR, "script_contra.jsx")
JSX_SCRIPT_PRODUTO = os.path.join(DATA_DIR, "script_produto.jsx")

PDF_PATH = os.path.join(DATA_DIR, "output", "resultado.pdf")

#-------------------------------------------------


# FUNÇÃO DE LOG DE AÇÕES
#-------------------------------------------------

def registrar_acao(user_code, acao):
    name = session.get("usuario")
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()
        query = """
            INSERT INTO historico (user_code, acao, data_hora, nome_usuario)
            VALUES (%s, %s, NOW(), %s)
        """
        cursor.execute(query, (user_code, acao, name))
        conn.commit()
    except Exception as e:
        print(f"Erro ao registrar ação: {e}")
    finally:
        cursor.close()
        conn.close()

#-------------------------------------------------


# ROUTES LOGIN/LOGOUT e SESSION
#-------------------------------------------------

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "usuario" not in session:
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function

@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        usuario = request.form.get("username")
        senha = request.form.get("password")
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor(dictionary=True)
        cursor.execute("SELECT * FROM usuarios WHERE user = %s", (usuario,))
        user = cursor.fetchone()

        cursor.close()
        conn.close()

        if not user:
            flash("Usuário ou senha incorretos.", "erro")
            return redirect(url_for("login"))

        if not check_password_hash(user["password_hash"], senha):
            flash("Usuário ou senha incorretos.", "erro")
            return redirect(url_for("login"))

        session["user_id"] = user["id"]
        session["usuario"] = user["user"]
        registrar_acao(session["user_id"], "Logou")

        flash("Login realizado com sucesso!", "sucesso")
        return redirect(url_for("index"))

    return render_template("login.html")

@app.route("/logout")
def logout():
    session.pop("usuario", None)
    return redirect(url_for("login"))

def check_session_queue(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        current_user = session.get("user_id")
        now = time.time()
        
        if SESSION_LOCK["user_id"] is not None:
            if now - SESSION_LOCK["last_active"] > LOCK_TIMEOUT:
                SESSION_LOCK["user_id"] = None
        
        if SESSION_LOCK["user_id"] is None or SESSION_LOCK["user_id"] == current_user:
            SESSION_LOCK["user_id"] = current_user
            SESSION_LOCK["last_active"] = now 
            return f(*args, **kwargs)
        
        else:
            return render_template("waiting.html")
            
    return decorated_function

@app.route("/liberar_sessao")
@login_required
def liberar_sessao():
    if SESSION_LOCK["user_id"] == session.get("user_id"):
        SESSION_LOCK["user_id"] = None
    print(f"🔓 Sessão liberada pelo usuário {session.get('usuario')}")
    return redirect(url_for("index"))



#-------------------------------------------------




# TEMPLATES
#-------------------------------------------------

@app.route("/")
@login_required
def index():
    return render_template("index.html", usuario=session["usuario"])


@app.route("/visualizar")
@login_required
@check_session_queue
def visualizar():
    return render_template("visualizer.html", usuario=session["usuario"])


@app.route('/painel', methods=["GET", "POST"])
@login_required
@check_session_queue
def painel():
    if request.method == "POST":
        layout_recebido = request.form.get('layout_escolhido')
        session['layout_escolhido'] = layout_recebido
        print(f"Layout salvo na sessão: {layout_recebido}")
            
    return render_template("painel.html", usuario=session["usuario"])


@app.route("/opcoes", methods=["POST"])
@login_required
@check_session_queue
def opcoes():
    raw_referencias = request.form.get("referencias")
    session['referencias'] = raw_referencias

    print(f"🔄 Iniciando atualização para: {raw_referencias}")
    processar_lista_referencias(raw_referencias)
    print("✅ Atualização concluída. Renderizando template.")

    return render_template("option.html")

@app.route("/admin")
@login_required
def admin():
    usuario = session["usuario"]

    if usuario not in ["Ana Lu", "Yasmin"]:
        return render_template(
            "acesso_negado.html",
            mensagem="Você não tem permissão para acessar esta página."
        )
    return render_template("admpainel.html", usuario=session["usuario"])


@app.route("/agente")
@login_required
def agente():
    usuario = session["usuario"]

    if usuario not in ["Ana Lu", "Yasmin", "Renan", "José", ]:
        return render_template(
            "acesso_negado.html",
            mensagem="Você não tem permissão para acessar esta página."
        )
    return render_template("agentepainel.html", usuario=session["usuario"])

#-------------------------------------------------




# FILE ROUTES
#-------------------------------------------------

@app.route('/foto/<ref>')
@login_required
def foto(ref):
    caminho = f"C:/Users/Administrador/Documents/fotosref/{ref}.jpg"
    return send_file(caminho)

@app.route('/resultado')
@login_required
def resultado():
    path = "C:\\Users\\Administrador\\Documents\\Sistemas\\PDFgenerator\\indesign\\output\\resultado.pdf"
    return send_file(path)

#-------------------------------------------------




# FUNÇÕES AUXILIARES
#-------------------------------------------------

def obter_slug_por_code(code):
    slug = None
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()
        cursor.execute("SELECT slug FROM products WHERE code = %s LIMIT 1", (code,))
        resultado = cursor.fetchone()
        if resultado:
            slug = resultado[0]
        cursor.close()
        conn.close()
    except Exception as e:
        print(f"⚠️ Erro ao buscar slug para o code '{code}': {e}")
    
    return slug

def worker_atualizar_ref(code):
    slug_db = obter_slug_por_code(code)
    
    if not slug_db:
        print(f"⏭️ Pular {code}: Slug não encontrado no banco de dados. {code}")
        return False

    url = (
        f"https://apivesti.vesti.mobi/appmarca/v1/products/company/"
        f"{COMPANY_ID}/product/{slug_db}/showcase?cid=7368dc35b43219a&reseller_id=null"
    )

    try:
        response = requests.get(url, timeout=10)
        response.raise_for_status()
        data = response.json().get("product_group", {})

        novo_price = data.get("price")
        novo_promotional = data.get("promotion")
        novo_price_promo = data.get("price_promotional")
        nova_compo = data.get("composition")
        product_id = data.get("id")
        
        sizes = data.get("sizes", [])
        sizes_names = ",".join(s["name"] for s in sizes)

        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()

        sql = """
            UPDATE products 
            SET 
                price = %s,
                promotional = %s,
                price_promotional = %s,
                composition = %s,
                sizes = %s,
                product_id = %s
            WHERE code = %s
        """
        
        vals = (
            novo_price, 
            novo_promotional, 
            novo_price_promo, 
            nova_compo, 
            sizes_names,
            product_id,
            code
        )

        cursor.execute(sql, vals)
        conn.commit()
        
        linhas_afetadas = cursor.rowcount
        cursor.close()
        conn.close()

        if linhas_afetadas > 0:
            print(f"✅ {code} (Slug: {slug_db}) atualizado com sucesso!")
        else:
            print(f"🔄️ {code} processado, mas banco não reportou mudanças (talvez dados iguais).")
            
        return True

    except Exception as e:
        print(f"❌ Erro ao atualizar {code} (via slug {slug_db}): {e}")
        return False

def processar_lista_referencias(referencias_str):
    if not referencias_str:
        return

    lista_codes = [
        code.strip() 
        for code in referencias_str.replace('\n', ',').split(',') 
        if code.strip()
    ]

    with ThreadPoolExecutor(max_workers=5) as executor:
        futures = [executor.submit(worker_atualizar_ref, code) for code in lista_codes]
        for future in as_completed(futures):
            future.result()

def kill_indesign_force():
    print("⚠️  Iniciando protocolo de encerramento forçado do InDesign...")
    try:
        subprocess.run(["taskkill", "/F", "/IM", "InDesign.exe"], 
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        subprocess.run(["taskkill", "/F", "/IM", "wscript.exe"], 
                    stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        print("✅  Processos do InDesign encerrados.")
    except Exception as e:
        print(f"❌  Erro ao tentar matar processos: {e}")


def escolher_script(capa: bool, contracapa: bool) -> str:
    if capa and contracapa:
        return JSX_SCRIPT_COMPLETO
    if capa and not contracapa:
        return JSX_SCRIPT_CAPA
    if not capa and contracapa:
        return JSX_SCRIPT_CONTRACAPA
    return JSX_SCRIPT_PRODUTO

def executar_indesign_with_jsx(jsx_path: str) -> bool:
    jsx_path = os.path.abspath(jsx_path)

    vbs_template = f'''
Set app = CreateObject("InDesign.Application")
app.DoScript "{jsx_path}", 1246973031
'''
    fd, vbs_temp_path = tempfile.mkstemp(suffix=".vbs", prefix="run_jsx_")
    os.close(fd)
    try:
        with open(vbs_temp_path, "w", encoding="utf-8") as f:
            f.write(vbs_template)

        resultado = subprocess.run(["wscript", vbs_temp_path], shell=False, capture_output=True, text=True, timeout=1800)
        if resultado.returncode == 0:
            print("InDesign executado com sucesso")
            return True
        else:
            print("Erro na execução do VBS:", resultado.stderr)
            return False
    except subprocess.TimeoutExpired:
        print("Execução do InDesign/VBS expirou.")
        kill_indesign_force()
        return False
    except Exception as e:
        print("Erro ao executar vbs temporário:", e)
        kill_indesign_force()
        return False
    finally:
        try:
            os.remove(vbs_temp_path)
        except Exception:
            pass

def clean_composition(text):
    if not text or pd.isna(text):
        return ""
    text = str(text)
    text = text.replace('\r', ' ').replace('\n', ' ').replace('\t', ' ')
    return text

def processar_geracao(dados_form, session_data):
    referencias = session_data.get("referencias")
    if not referencias:
        return False

    if isinstance(referencias, str):
        if "," in referencias:
            referencias = [r.strip() for r in referencias.split(",") if r.strip()]
        else:
            referencias = [r.strip() for r in referencias.split() if r.strip()]

    conn = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor(buffered=True)

    query = """
        SELECT name, price, promotional, price_promotional, composition, sizes
        FROM products WHERE code = %s
    """

    want_referencia = bool(dados_form.get("referencia", False))
    want_preco = bool(dados_form.get("preco", False))
    want_composicao = bool(dados_form.get("composicao", False))
    want_tamanho = bool(dados_form.get("tamanho", False))
    want_capa = bool(dados_form.get("capa", False))
    want_contracapa = bool(dados_form.get("contracapa", False))
    want_logo = bool(dados_form.get("logo", False))
    want_sublogo = bool(dados_form.get("sublogo", False))
    want_mensagem = bool(dados_form.get("mensagem", False))
    referencia_capa_val = dados_form.get("referenciaCapa", "")
    logo_escolhida = dados_form.get("logoescolhida", "")
    sublogo_escolhida = dados_form.get("sublogoescolhida", "")
    mensagem_escolhida = dados_form.get("mensagemescolhida", "")
    print(f"Mensagem escolhida: {mensagem_escolhida}")

    lista_produtos = []
    for ref in referencias:
        cursor.execute(query, (ref,))
        result = cursor.fetchone()
        if not result: continue

        name, price, promotional, price_promotional, composition, sizes_raw = result
        tms = {t: "" for t in ["pp", "p", "m", "g", "gg", "u"]}
        pre = {t: "" for t in ["pp", "p", "m", "g", "gg", "u"]}

        if want_tamanho:
            texto_tamanho = "Tamanhos disponíveis:"
            circulo = "O"
            fixed_pp, fixed_p, fixed_m = "PP", "P", "M"
            fixed_g, fixed_gg, fixed_u = "G", "GG", "U"
            if sizes_raw:
                lista_tamanhos_db = [s.strip().upper() for s in sizes_raw.split(",")]
                for tam in tms.keys():
                    if tam.upper() in lista_tamanhos_db:
                        tms[tam] = tam.upper()
                        pre[tam] = "l"
        else:
            texto_tamanho = circulo = ""
            fixed_pp = fixed_p = fixed_m = fixed_g = fixed_gg = fixed_u = ""

        if want_preco:
            if promotional == 1:
                de_ou_por, preco_original, traco, real, preco_promocional, por, realfixo = "DE:", price, "\\", "R$", price_promotional, "POR:", "R$"
            else:
                de_ou_por, preco_original, traco, real, preco_promocional, por, realfixo = "POR:", price, "", "", "", "", "R$"
        else:
            de_ou_por = preco_original = traco = real = preco_promocional = por = realfixo = ""

        comp = ""
        if want_composicao:
            comp_bruta = clean_composition(composition)
            comp = (comp_bruta[:28] + "+") if len(comp_bruta) > 28 else comp_bruta

        foto_path = rf"C:\Users\Administrador\Documents\fotosref\{ref}.jpg"
        lista_produtos.append({
            "referencia": ref if want_referencia else "",
            "nome": name if want_referencia else "",
            "de_ou_por": de_ou_por, "preco_original": preco_original,
            "traco": traco, "por": por, "real": real, "realfixo": realfixo,
            "preco_promocional": preco_promocional, "composicao": comp,
            "texto_tamanho": texto_tamanho, "circulo": circulo,
            "fixed_pp": fixed_pp, "tamanho PP": tms["pp"], "preenchimento PP": pre["pp"],
            "fixed_p": fixed_p, "tamanho P": tms["p"], "preenchimento P": pre["p"],
            "fixed_m": fixed_m, "tamanho M": tms["m"], "preenchimento M": pre["m"],
            "fixed_g": fixed_g, "tamanho G": tms["g"], "preenchimento G": pre["g"],
            "fixed_gg": fixed_gg, "tamanho GG": tms["gg"], "preenchimento GG": pre["gg"],
            "fixed_u": fixed_u, "tamanho U": tms["u"], "preenchimento U": pre["u"],
            "@fotos": foto_path
        })

    cursor.close()
    conn.close()

    if not lista_produtos:
        return False

    col_order = [
        "referencia", "nome", "de_ou_por", "preco_original", "traco", "por", "real", "realfixo",
        "preco_promocional", "composicao", "texto_tamanho", "circulo",
        "fixed_pp", "tamanho PP", "preenchimento PP",
        "fixed_p", "tamanho P", "preenchimento P",
        "fixed_m", "tamanho M", "preenchimento M",
        "fixed_g", "tamanho G", "preenchimento G",
        "fixed_gg", "tamanho GG", "preenchimento GG",
        "fixed_u", "tamanho U", "preenchimento U", "@fotos"
    ]
    pd.DataFrame(lista_produtos).reindex(columns=col_order).to_csv(
        CSV_PRODUTO_PATH, index=False, sep=";", encoding="utf-16"
    )

    # Capa
    if want_capa:
        pd.DataFrame([{
            "@fotofundo": rf"C:\Users\Administrador\Documents\fotosref\{referencia_capa_val}.jpg" if referencia_capa_val else "",
            "@logo": rf"C:\Users\Administrador\Documents\Sistemas\PDFgenerator\static\logos\{logo_escolhida}.png" if want_logo and logo_escolhida else "",
            "@sublogo": rf"C:\Users\Administrador\Documents\Sistemas\PDFgenerator\static\logos\{sublogo_escolhida}.png" if want_sublogo and sublogo_escolhida else "",
            "@mensagem": rf"C:\Users\Administrador\Documents\Sistemas\PDFgenerator\static\images\{mensagem_escolhida}.png" if want_mensagem and mensagem_escolhida else ""
        }]).to_csv(CSV_CAPA_PATH, index=False, sep=",", encoding="utf-16")
    elif os.path.exists(CSV_CAPA_PATH):
        os.remove(CSV_CAPA_PATH)

    # Contracapa
    if want_contracapa:
        pd.DataFrame([{
            "@fotofundo": rf"C:\Users\Administrador\Documents\fotosref\{referencia_capa_val}.jpg" if referencia_capa_val else "",
            "@logo": rf"C:\Users\Administrador\Documents\Sistemas\PDFgenerator\static\logos\{logo_escolhida}.png" if want_logo and logo_escolhida else "",
            "@sublogo": rf"C:\Users\Administrador\Documents\Sistemas\PDFgenerator\static\logos\{sublogo_escolhida}.png" if want_sublogo and sublogo_escolhida else ""
        }]).to_csv(CSV_CONTRACAPA_PATH, index=False, sep=",", encoding="utf-16")
    elif os.path.exists(CSV_CONTRACAPA_PATH):
        os.remove(CSV_CONTRACAPA_PATH)

    # Layout e InDesign
    with open(os.path.join(DATA_DIR, "layout_config.txt"), "w") as f:
        f.write(f"{session_data.get('layout_escolhido')}.indd")

    sucesso = executar_indesign_with_jsx(escolher_script(want_capa, want_contracapa))
    
    try:
        conn2 = mysql.connector.connect(**DB_CONFIG)
        cursor2 = conn2.cursor()
        cursor2.execute(
            "INSERT INTO historico (user_code, acao, data_hora, nome_usuario) VALUES (%s, %s, NOW(), %s)",
            (session_data["user_id"], "Criou PDF", session_data["usuario"])
        )
        conn2.commit()
        cursor2.close()
        conn2.close()
    except Exception as e:
        print(f"Erro ao registrar ação: {e}")

    return sucesso

JOBS = {}

# CADASTRAR
CADASTRO_JOBS = {}


def _cadastrar_worker(job_id, codes):
    logs = CADASTRO_JOBS[job_id]["logs"]

    for code in codes:
        url = (
            f"https://apivesti.vesti.mobi/appmarca/v1/products/company/"
            f"{COMPANY_ID}/product/{code}/showcase?cid=7368dc35b43219a&reseller_id=null"
        )
        try:
            r = requests.get(url, timeout=10)
            r.raise_for_status()
            data = r.json()["product_group"]
            sizes_names = ",".join(s["name"] for s in data.get("sizes", []))

            conn = mysql.connector.connect(**DB_CONFIG)
            cursor = conn.cursor()
            cursor.execute("""
                INSERT INTO products (code, name, slug, price, promotional, price_promotional, composition, product_id, sizes)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s)
                ON DUPLICATE KEY UPDATE name=%s,slug=%s,price=%s,promotional=%s,price_promotional=%s,composition=%s,product_id=%s,sizes=%s
            """, (
                data['code'], data['name'], data['slug'], data['price'], data['promotion'],
                data['price_promotional'], data['composition'], data['id'], sizes_names,
                data['name'], data['slug'], data['price'], data['promotion'],
                data['price_promotional'], data['composition'], data['id'], sizes_names
            ))
            conn.commit()
            cursor.close()
            conn.close()
            logs.append(f"✅ {code} cadastrado")

            product_id = data['id']
            product_code = data['code']

            url_v2 = f"{API_URL_BASE}/v2/product/company/{COMPANY_ID}/product/{product_id}"
            headers = {"apikey": os.getenv('API_KEY')}
            r2 = requests.get(url_v2, headers=headers, timeout=5)
            colors = r2.json().get("response", {}).get("colors", [])

            if colors:
                conn2 = mysql.connector.connect(**DB_CONFIG)
                cursor2 = conn2.cursor()
                for cor in colors:
                    cursor2.execute("""
                        INSERT INTO product_colors (product_id, product_code, color_name, color_code)
                        VALUES (%s,%s,%s,%s)
                        ON DUPLICATE KEY UPDATE color_name=VALUES(color_name)
                    """, (product_id, product_code, cor["name"], cor["code"]))
                conn2.commit()
                cursor2.close()
                conn2.close()
                logs.append(f"🎨 Cores de {code} sincronizadas ({len(colors)})")

        except Exception as e:
            logs.append(f"❌ Erro em {code}: {e}")

    CADASTRO_JOBS[job_id]["status"] = "concluido"


@app.route("/cadastrarproduto", methods=["POST"])
@login_required
def cadastrarproduto():
    raw = request.form.get("cadastrar", "")
    codes = [c.strip() for c in raw.replace('\n', ' ').split() if c.strip()]
    if not codes:
        return jsonify({"erro": "Nenhum código enviado."}), 400

    job_id = str(uuid_lib.uuid4())
    CADASTRO_JOBS[job_id] = {"status": "rodando", "logs": []}

    threading.Thread(target=_cadastrar_worker, args=(job_id, codes), daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/status_cadastro/<job_id>")
@login_required
def status_cadastro(job_id):
    job = CADASTRO_JOBS.get(job_id)
    if not job:
        return jsonify({"status": "não encontrado"}), 404
    return jsonify(job)

# VERIFICAR

VERIFICA_JOBS = {}

def _verificar_worker(job_id, codes, tipo):
    logs = VERIFICA_JOBS[job_id]["logs"]
    conn = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor(dictionary=True)

    for code in codes:
        try:
            if tipo == "slug":
                cursor.execute("SELECT * FROM products WHERE slug = %s LIMIT 1", (code,))
            else:
                cursor.execute("SELECT * FROM products WHERE product_id = %s LIMIT 1", (code,))

            row = cursor.fetchone()

            if not row:
                logs.append(f"❌ Referência não encontrada no banco de dados: {code}")
                continue

            ref_display = row.get("code", code)

            if tipo == "slug":
                url = (
                    f"https://apivesti.vesti.mobi/appmarca/v1/products/company/"
                    f"{COMPANY_ID}/product/{row['slug']}/showcase?cid=7368dc35b43219a&reseller_id=null"
                )
                r = requests.get(url, timeout=10)
                r.raise_for_status()
                api_data = r.json()["product_group"]

                atualizado = (
                    float(row.get("price") or 0) == float(api_data.get("price") or 0) and
                    float(row.get("price_promotional") or 0) == float(api_data.get("price_promotional") or 0) and
                    str(row.get("name") or "").strip() == str(api_data.get("name") or "").strip()
                )

            else:
                url_v2 = f"{API_URL_BASE}/v2/product/company/{COMPANY_ID}/product/{row['product_id']}"
                headers = {"apikey": APIKEY}
                r = requests.get(url_v2, headers=headers, timeout=10)
                r.raise_for_status()
                api_data = r.json().get("response", {})

                atualizado = (
                    float(row.get("price") or 0) == float(api_data.get("price") or 0) and
                    float(row.get("price_promotional") or 0) == float(api_data.get("price_promotional") or 0) and
                    str(row.get("name") or "").strip() == str(api_data.get("name") or "").strip()
                )

            if atualizado:
                logs.append(f"✅ Referência existe e atualizada: {ref_display}")
            else:
                logs.append(f"Nova consulta API para {ref_display} retornou: price={api_data.get('price')}, price_promotional={api_data.get('price_promotional')}, name={api_data.get('name')}")
                logs.append(f"⚠️ Referência existe mas desatualizada: {ref_display}")

        except Exception as e:
            logs.append(f"❌ Erro ao verificar {code}: {e}")

    cursor.close()
    conn.close()
    VERIFICA_JOBS[job_id]["status"] = "concluido"


@app.route("/verificaproduto", methods=["POST"])
@login_required
def verificaproduto():
    raw = request.form.get("verificar", "")
    tipo = request.form.get("tipo_verifica", "slug")
    codes = [c.strip() for c in raw.replace('\n', ' ').split() if c.strip()]
    if not codes:
        return jsonify({"erro": "Nenhum código enviado."}), 400

    job_id = str(uuid_lib.uuid4())
    VERIFICA_JOBS[job_id] = {"status": "rodando", "logs": []}

    threading.Thread(target=_verificar_worker, args=(job_id, codes, tipo), daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/status_verifica/<job_id>")
@login_required
def status_verifica(job_id):
    job = VERIFICA_JOBS.get(job_id)
    if not job:
        return jsonify({"status": "não encontrado"}), 404
    return jsonify(job)

#-------------------------------------------------




# ROUTES PRINCIPAIS DE GERAÇÃO
#-------------------------------------------------


@app.route("/gerar_planilha", methods=["POST"])
@login_required
@check_session_queue
def gerar_planilha():
    dados_json = request.form.get("dados_json")
    if not dados_json:
        return jsonify({"erro": "dados_json não enviado."}), 400

    try:
        dados_form = json.loads(dados_json)
    except Exception as e:
        return jsonify({"erro": "dados_json inválido.", "detalhe": str(e)}), 400

    nome_do_arquivo = dados_form.get("nomeArquivo", "arquivo")
    session["nome_arquivo_escolhido"] = nome_do_arquivo

    job_id = str(uuid_lib.uuid4())
    JOBS[job_id] = {"status": "rodando"}

    session_data = {
        "user_id": session.get("user_id"),
        "usuario": session.get("usuario"),
        "referencias": session.get("referencias"),
        "layout_escolhido": session.get("layout_escolhido"),
        "nome_arquivo_escolhido": nome_do_arquivo,
    }

    def rodar_job():
        try:
            resultado = processar_geracao(dados_form, session_data)
            JOBS[job_id]["status"] = "ok" if resultado else "erro"
        except Exception as e:
            print(f"Erro no job: {e}")
            JOBS[job_id]["status"] = "erro"

    threading.Thread(target=rodar_job, daemon=True).start()
    return jsonify({"job_id": job_id})


@app.route("/status_job/<job_id>")
@login_required
def status_job(job_id):
    job = JOBS.get(job_id)
    if not job:
        return jsonify({"status": "não encontrado"}), 404
    return jsonify(job)


@app.route("/download")
@login_required
def download_pdf():
    if not os.path.exists(PDF_PATH):
        return jsonify({"erro": "PDF não encontrado."}), 404

    nome = session.get("nome_arquivo_escolhido")
    nome = "".join(c for c in nome if c.isalnum() or c in (" ", "-", "_")).strip()
    if not nome:
        nome = "arquivo_final"

    return send_file(
        PDF_PATH,
        as_attachment=True,
        download_name=f"{nome}.pdf"
    )

@app.route('/baixar_fotos')
@login_required
def baixar_fotos():
    path = r"C:\Users\Administrador\Documents\Sistemas\PDFgenerator\indesign\output\resultado.pdf"
    memory_zip = io.BytesIO()
    with zipfile.ZipFile(memory_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
        pdf_document = fitz.open(path)
        for page_num in range(len(pdf_document)):
            page = pdf_document.load_page(page_num)
            pix = page.get_pixmap(dpi=75)
            img_bytes = pix.tobytes("png")
            nome_arquivo = session.get("nome_arquivo_escolhido", "pagina") + f"_{page_num + 1}.png"
            
            zf.writestr(nome_arquivo, img_bytes)
    memory_zip.seek(0)
    return send_file(
        memory_zip,
        mimetype='application/zip',
        as_attachment=True,
        download_name='fotos_do_pdf.zip'
    )

@app.route("/listar_usuarios")
@login_required
def listar_usuarios():
    conn = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT id, user FROM usuarios ORDER BY user")
    users = cursor.fetchall()
    cursor.close()
    conn.close()
    return jsonify(users)


@app.route("/novo_usuario", methods=["POST"])
@login_required
def novo_usuario():
    username = request.form.get("newusername", "").strip()
    password = request.form.get("password", "")
    confirm  = request.form.get("confirm_password", "")

    if not username or not password:
        return jsonify({"erro": "Preencha todos os campos."}), 400
    if password != confirm:
        return jsonify({"erro": "Senhas não coincidem."}), 400

    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()
        cursor.execute("INSERT INTO usuarios (user, password_hash) VALUES (%s, %s)",
            (username, generate_password_hash(password)))
        conn.commit()
        cursor.close()
        conn.close()
        registrar_acao(session["user_id"], f"Criou usuário: {username}")
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


@app.route("/editar_usuario", methods=["POST"])
@login_required
def editar_usuario():
    user_id  = request.form.get("user_id")
    password = request.form.get("newpassword", "")
    confirm  = request.form.get("confirm_newpassword", "")

    if not password:
        return jsonify({"erro": "Informe a nova senha."}), 400
    if password != confirm:
        return jsonify({"erro": "Senhas não coincidem."}), 400

    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()
        cursor.execute("UPDATE usuarios SET password_hash = %s WHERE id = %s",
            (generate_password_hash(password), user_id))
        conn.commit()
        cursor.close()
        conn.close()
        registrar_acao(session["user_id"], f"Editou senha do usuário ID: {user_id}")
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


@app.route("/excluir_usuario", methods=["POST"])
@login_required
def excluir_usuario():
    user_id = request.form.get("user_id")
    try:
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor()
        cursor.execute("DELETE FROM historico WHERE user_code = %s", (user_id,))
        cursor.execute("DELETE FROM usuarios WHERE id = %s", (user_id,))
        conn.commit()
        cursor.close()
        conn.close()
        registrar_acao(session["user_id"], f"Excluiu usuário ID: {user_id}")
        return jsonify({"ok": True})
    except Exception as e:
        return jsonify({"erro": str(e)}), 500


@app.route("/historico_usuario/<int:user_id>")
@login_required
def historico_usuario(user_id):
    dias = request.args.get("dias", "hoje")
    conn = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor(dictionary=True)

    if dias == "hoje":
        cursor.execute("""
            SELECT acao, data_hora FROM historico
            WHERE user_code = %s AND DATE(data_hora) = CURDATE()
            ORDER BY data_hora DESC
        """, (user_id,))
    else:
        cursor.execute("""
            SELECT acao, data_hora FROM historico
            WHERE user_code = %s AND data_hora >= DATE_SUB(NOW(), INTERVAL 15 DAY)
            ORDER BY data_hora DESC
        """, (user_id,))

    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    for r in rows:
        r["data_hora"] = r["data_hora"].strftime("%d/%m/%Y %H:%M")
    return jsonify(rows)

#-------------------------------------------------


#----- AGENTES IA -----

@app.route('/api/articles')
@login_required
def get_articles():
    rows = sheet.worksheet('Artigos').get_all_records()
    return jsonify(rows)

@app.route('/api/articles/<article_id>', methods=['DELETE'])
@login_required
def delete_article(article_id):
    ws = sheet.worksheet('Artigos')
    cell = ws.find(article_id)
    if cell:
        ws.delete_rows(cell.row)
    return jsonify({'ok': True})

@app.route('/api/articles/<article_id>', methods=['PATCH'])
@login_required
def update_article(article_id):
    ws = sheet.worksheet('Artigos')
    cell = ws.find(article_id)
    if not cell:
        return jsonify({'ok': False}), 404
    
    data = request.json
    headers = ws.row_values(1)
    
    for field, value in data.items():
        if field in headers:
            col = headers.index(field) + 1
            ws.update_cell(cell.row, col, value)
    
    return jsonify({'ok': True})

@app.route('/api/pesquisar', methods=['POST'])
@login_required
@check_session_queue
def pesquisar():
    data = request.get_json()
    response = requests.post(
        'http://localhost:5678/webhook/pesquisador',
        json=data,
        timeout=120
    )
    print("RESPOSTA N8N:", response.text[:200])
    print("RESPOSTA COMPLETA:", repr(response.text))
    import json
    text = response.text.strip().lstrip('=')
    return jsonify(json.loads(text))

@app.route('/api/temas')
@login_required
def get_temas():
    ws = sheet.worksheet('Temasnovos')
    # linhas 2 a 6, colunas A (slug) e B (titulo)
    slugs  = ws.col_values(1)[1:6]   # coluna A, linhas 2-6
    titulos = ws.col_values(2)[1:6]  # coluna B, linhas 2-6
    
    temas = [
        { "slug": slugs[i], "titulo": titulos[i] }
        for i in range(len(slugs))
        if slugs[i] and titulos[i]
    ]
    return jsonify(temas)

@app.route('/api/angulador', methods=['POST'])
@login_required
@check_session_queue
def angulador():
    data = request.get_json()
    response = requests.post(
        'http://localhost:5678/webhook/angulador',
        json=data,
        timeout=120
    )
    print("STATUS:", response.status_code)
    print("RESPOSTA ANGULADOR:", response.text[:200])
    
    import json
    text = response.text.strip().lstrip('=')
    parsed = json.loads(text)

    if isinstance(parsed, list):
        return jsonify({ "angulos": parsed })
    
    return jsonify(parsed)

@app.route('/api/gerar', methods=['POST'])
@login_required
@check_session_queue
def gerar():
    data = request.get_json()
    response = requests.post(
        'http://localhost:5678/webhook/redator',
        json=data,
        timeout=300  # artigo demora mais
    )
    print("RESPOSTA REDATOR:", response.text[:200])
    text = response.text.strip().lstrip('=')
    import json
    return jsonify(json.loads(text))

@app.route('/api/roteiro', methods=['POST'])
@login_required
@check_session_queue
def roteiro():
    data = request.get_json()
    response = requests.post(
        'http://localhost:5678/webhook/marketing',
        json={
            'concorrente': data.get('concorrente'),
            'max_results': 20
        },
        timeout=300
    )
    print("RESPOSTA ROTEIRO:", response.text[:200])
    text = response.text.strip().lstrip('=')
    import json
    text = response.text.strip().lstrip('=')
    return jsonify(json.loads(text))
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)