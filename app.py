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


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "indesign")
CSV_DIR = os.path.join(DATA_DIR, "CSV")
os.makedirs(CSV_DIR, exist_ok=True)

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
    layout_escolhido = request.form.get('layout_escolhido')
    print(f"Layout escolhido: {layout_escolhido}")
    session['layout_escolhido'] = layout_escolhido
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

#-------------------------------------------------




# FILE ROUTES
#-------------------------------------------------

@app.route('/foto/<ref>')
def foto(ref):
    caminho = f"C:/Users/Administrador/Documents/fotosref/{ref}.jpg"
    return send_file(caminho)

@app.route('/resultado')
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
        print(f"⏭️ Pular {code}: Slug não encontrado no banco de dados.")
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
            print(f"⚠️ {code} processado, mas banco não reportou mudanças (talvez dados iguais).")
            
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

#-------------------------------------------------




# ROUTES PRINCIPAIS DE GERAÇÃO
#-------------------------------------------------

@app.route("/gerar_planilha", methods=["POST"])
def gerar_planilha():
    dados_json = request.form.get("dados_json")

    if not dados_json:
        return jsonify({"erro": "dados_json não enviado."}), 400

    try:
        dados_form = json.loads(dados_json)
    except Exception as e:
        return jsonify({"erro": "dados_json inválido.", "detalhe": str(e)}), 400

    session["nome_arquivo_escolhido"] = dados_form.get("nomeArquivo", "arquivo")
    referencias = session.get("referencias")
    if not referencias:
        return jsonify({"erro": "Nenhuma referência salva na sessão."}), 400

    if isinstance(referencias, str):
        if "," in referencias:
            referencias = [r.strip() for r in referencias.split(",") if r.strip()]
        else:
            referencias = [r.strip() for r in referencias.split() if r.strip()]
    elif isinstance(referencias, list):
        pass
    else:
        try:
            referencias = json.loads(referencias)
            if not isinstance(referencias, list): raise ValueError
        except Exception:
            return jsonify({"erro": "Formato de referências inválido."}), 400

    conn = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor(buffered=True)

    query = """
        SELECT name, price, promotional, price_promotional, composition, sizes
        FROM products
        WHERE code = %s
    """

    lista_produtos = []
    want_referencia = bool(dados_form.get("referencia", False))
    want_preco = bool(dados_form.get("preco", False))
    want_composicao = bool(dados_form.get("composicao", False))
    want_tamanho = bool(dados_form.get("tamanho", False))
    want_capa = bool(dados_form.get("capa", False))
    want_contracapa = bool(dados_form.get("contracapa", False))
    want_logo = bool(dados_form.get("logo", False))
    want_sublogo = bool(dados_form.get("sublogo", False))
    referencia_capa_val = dados_form.get("referenciaCapa", "")
    logo_escolhida = dados_form.get("logoescolhida", "")
    sublogo_escolhida = dados_form.get("sublogoescolhida", "")

    for ref in referencias:
        cursor.execute(query, (ref,))
        result = cursor.fetchone()
        if not result: continue

        name, price, promotional, price_promotional, composition, sizes_raw = result

        tms = {f"{t}": "" for t in ["pp", "p", "m", "g", "gg", "u"]}
        pre = {f"{t}": "" for t in ["pp", "p", "m", "g", "gg", "u"]}
        
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

        linha_produto = {
            "referencia": ref if want_referencia else "",
            "nome": name if want_referencia else "",
            "de_ou_por": de_ou_por,
            "preco_original": preco_original,
            "traco": traco,
            "por": por,
            "real": real,
            "realfixo": realfixo,
            "preco_promocional": preco_promocional,
            "composicao": comp,
            "texto_tamanho": texto_tamanho, 
            "circulo": circulo,             
            "fixed_pp": fixed_pp,           
            "tamanho PP": tms["pp"],        
            "preenchimento PP": pre["pp"],  
            "fixed_p": fixed_p,             
            "tamanho P": tms["p"],          
            "preenchimento P": pre["p"],    
            "fixed_m": fixed_m,             
            "tamanho M": tms["m"],          
            "preenchimento M": pre["m"],    
            "fixed_g": fixed_g,             
            "tamanho G": tms["g"],          
            "preenchimento G": pre["g"],    
            "fixed_gg": fixed_gg,           
            "tamanho GG": tms["gg"],        
            "preenchimento GG": pre["gg"],  
            "fixed_u": fixed_u,             
            "tamanho U": tms["u"],          
            "preenchimento U": pre["u"],    
            "@fotos": foto_path             
        }
        lista_produtos.append(linha_produto)

    cursor.close()
    conn.close()

    if not lista_produtos:
        return jsonify({"erro": "Nenhuma referência encontrada."}), 404

    df_produtos = pd.DataFrame(lista_produtos)
    
    col_order = [
        "referencia", "nome", "de_ou_por", "preco_original", "traco", "por", "real", "realfixo", "preco_promocional", "composicao",
        "texto_tamanho", "circulo", 
        "fixed_pp", "tamanho PP", "preenchimento PP",
        "fixed_p", "tamanho P", "preenchimento P",
        "fixed_m", "tamanho M", "preenchimento M",
        "fixed_g", "tamanho G", "preenchimento G",
        "fixed_gg", "tamanho GG", "preenchimento GG",
        "fixed_u", "tamanho U", "preenchimento U",
        "@fotos"
    ]
    
    df_produtos = df_produtos.reindex(columns=col_order)
    df_produtos.to_csv(CSV_PRODUTO_PATH, index=False, sep=";", encoding="utf-16")


    if want_capa:
        linha_capa = {
            "@fotofundo": rf"C:\Users\Administrador\Documents\fotosref\{referencia_capa_val}.jpg" if referencia_capa_val else "",
            "@logo": rf"C:\Users\Administrador\Documents\Sistemas\PDFgenerator\static\logos\{logo_escolhida}.png" if want_logo and logo_escolhida else "",
            "@sublogo": rf"C:\Users\Administrador\Documents\Sistemas\PDFgenerator\static\logos\{sublogo_escolhida}.png" if want_sublogo and sublogo_escolhida else ""
        }
        df_capa = pd.DataFrame([linha_capa])
        df_capa.to_csv(CSV_CAPA_PATH, index=False, sep=",", encoding="utf-16")
    else:
        if os.path.exists(CSV_CAPA_PATH):
            os.remove(CSV_CAPA_PATH)

    if want_contracapa:
        linha_contracapa = {
            "@fotofundo": rf"C:\Users\Administrador\Documents\fotosref\{referencia_capa_val}.jpg" if referencia_capa_val else "",
            "@logo": rf"C:\Users\Administrador\Documents\Sistemas\PDFgenerator\static\logos\{logo_escolhida}.png" if want_logo and logo_escolhida else "",
            "@sublogo": rf"C:\Users\Administrador\Documents\Sistemas\PDFgenerator\static\logos\{sublogo_escolhida}.png" if want_sublogo and sublogo_escolhida else ""
        }
        df_contra = pd.DataFrame([linha_contracapa])
        df_contra.to_csv(CSV_CONTRACAPA_PATH, index=False, sep=",", encoding="utf-16")
    else:
        if os.path.exists(CSV_CONTRACAPA_PATH):
            os.remove(CSV_CONTRACAPA_PATH)

    script_to_run = escolher_script(want_capa, want_contracapa)

    sucesso = executar_indesign_with_jsx(script_to_run)
    registrar_acao(session["user_id"], "Criou PDF: ".format(session.get("nome_arquivo_escolhido")))

    if sucesso:
        time.sleep(5)
        return redirect(url_for('visualizar'))
    else:
        flash("Ocorreu um erro crítico no InDesign ou Timeout. O processo foi reiniciado por segurança.", "erro")
        return redirect(url_for('liberar_sessao'))
    


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

#-------------------------------------------------


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)