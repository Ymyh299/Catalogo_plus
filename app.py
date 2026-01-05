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

app = Flask(__name__)
app.secret_key = "anag25000"

DB_CONFIG = {
    "host": "localhost",
    "user": "root",
    "password": "QualquerUma123",
    "database": "catalogoplus"
}

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

        # Conectar ao banco
        conn = mysql.connector.connect(**DB_CONFIG)
        cursor = conn.cursor(dictionary=True)

        cursor.execute("SELECT * FROM usuarios WHERE user = %s", (usuario,))
        user = cursor.fetchone()

        cursor.close()
        conn.close()

        if not user:
            flash("Usuário ou senha incorretos.", "erro")
            return redirect(url_for("login"))

        # Verificar senha usando bcrypt (Werkzeug)
        if not check_password_hash(user["password_hash"], senha):
            flash("Usuário ou senha incorretos.", "erro")
            return redirect(url_for("login"))

        # Login OK
        session["user_id"] = user["id"]
        session["usuario"] = user["user"]

        flash("Login realizado com sucesso!", "sucesso")
        return redirect(url_for("index"))

    return render_template("login.html")


@app.route("/")
@login_required
def index():
    return render_template("index.html", usuario=session["usuario"])

# @app.route("/painel")
# @login_required
# def painel():
#     return render_template("painel.html", usuario=session["usuario"])




@app.route("/visualizar")
@login_required
def visualizar():
    return render_template("visualizer.html", usuario=session["usuario"])




@app.route('/foto/<ref>')
def foto(ref):
    caminho = f"C:/Users/Administrador/Documents/fotosref/{ref}.jpg"
    return send_file(caminho)

@app.route('/resultado')
def resultado():
    path = "C:\\Users\\Administrador\\Documents\\Sistemas\\PDFgenerator\\indesign\\output\\resultado.pdf"
    return send_file(path)




@app.route('/painel', methods=["GET", "POST"])
@login_required
def painel():
    layout_escolhido = request.form.get('layout_escolhido')
    print(f"Layout escolhido: {layout_escolhido}")
    session['layout_escolhido'] = layout_escolhido
    return render_template("painel.html", usuario=session["usuario"])

    

@app.route("/opcoes", methods=["POST"])
@login_required
def opcoes():
    session['referencias'] = request.form.get("referencias")
    return render_template("option.html")





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
    # criar arquivo VBS temporário
    fd, vbs_temp_path = tempfile.mkstemp(suffix=".vbs", prefix="run_jsx_")
    os.close(fd)
    try:
        with open(vbs_temp_path, "w", encoding="utf-8") as f:
            f.write(vbs_template)

        resultado = subprocess.run(["wscript", vbs_temp_path], shell=False, capture_output=True, text=True, timeout=120)
        if resultado.returncode == 0:
            print("InDesign executado com sucesso")
            return True
        else:
            print("Erro na execução do VBS:", resultado.stderr)
            return False
    except subprocess.TimeoutExpired:
        print("Execução do InDesign/VBS expirou.")
        return False
    except Exception as e:
        print("Erro ao executar vbs temporário:", e)
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

@app.route("/gerar_planilha", methods=["POST"])
def gerar_planilha():
    # receber dados do form (JSON string)
    dados_json = request.form.get("dados_json")




    if not dados_json:
        return jsonify({"erro": "dados_json não enviado."}), 400

    try:
        dados_form = json.loads(dados_json)
    except Exception as e:
        return jsonify({"erro": "dados_json inválido.", "detalhe": str(e)}), 400


    session["nome_arquivo_escolhido"] = dados_form.get("nomeArquivo", "arquivo")
    # pegar referências da sessão
    referencias = session.get("referencias")
    if not referencias:
        return jsonify({"erro": "Nenhuma referência salva na sessão."}), 400

    # aceitar string separada por vírgulas ou espaços, ou lista já serializada
    if isinstance(referencias, str):
        # tenta detectar separador
        if "," in referencias:
            referencias = [r.strip() for r in referencias.split(",") if r.strip()]
        else:
            referencias = [r.strip() for r in referencias.split() if r.strip()]
    elif isinstance(referencias, list):
        pass
    else:
        # se veio como JSON string dentro da session
        try:
            referencias = json.loads(referencias)
            if not isinstance(referencias, list):
                raise ValueError
        except Exception:
            return jsonify({"erro": "Formato de referências inválido."}), 400

    # conectar DB
    conn = mysql.connector.connect(**DB_CONFIG)
    cursor = conn.cursor(buffered=True)


    query = """
        SELECT name, price, promotional, price_promotional, composition
        FROM products
        WHERE code = %s
    """

    # listas / dicts para os 3 CSVs
    lista_produtos = []
    linha_capa = None
    linha_contracapa = None

    # flags do formulário (usar .get com default False para evitar KeyError)
    want_referencia = bool(dados_form.get("referencia", False))
    want_preco = bool(dados_form.get("preco", False))
    want_composicao = bool(dados_form.get("composicao", False))
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

        if not result:
            # se não achar no DB, pule
            continue

        name, price, promotional, price_promotional, composition = result

        # preço
        if want_preco:
            if promotional == 1:
                de_ou_por = "DE:"
                preco_original = price
                traco = "\\"
                real = "R$"
                preco_promocional = price_promotional
                por = "POR:"
                realfixo = "R$"
            else:
                de_ou_por = "POR:"
                realfixo = ""
                preco_original = price
                traco = ""
                real = ""
                preco_promocional = ""
                por = ""
        else:
            de_ou_por = ""
            preco_original = ""
            traco = ""
            real = ""
            preco_promocional = ""
            por = ""
            realfixo = "" 

        # composição
        comp = clean_composition(composition) if want_composicao else ""

        # foto (sempre obrigatória conforme você disse)
        foto_path = rf"C:\Users\Administrador\Documents\fotosref\{ref}.jpg"

        # monta linha produto (todas as colunas que pediu)
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
            "@fotos": foto_path
        }

        lista_produtos.append(linha_produto)

    cursor.close()
    conn.close()

    if not lista_produtos:
        return jsonify({"erro": "Nenhuma referência encontrada no banco."}), 404

    # montar CSV de produtos
    df_produtos = pd.DataFrame(lista_produtos)
    # garantir ordem de colunas (opcional)
    col_order = ["referencia", "nome", "de_ou_por", "preco_original", "traco", "por", "real", "realfixo", "preco_promocional", "composicao", "@fotos"]
    df_produtos = df_produtos.reindex(columns=col_order)
    df_produtos.to_csv(CSV_PRODUTO_PATH, index=False, sep=";", encoding="utf-16")

    # montar CSV de capa e contracapa (cada um tem apenas 1 linha com @fotofundo, @logo, @sublogo)
    # Se houver capa ativa, monta com os valores vindos do form
    if want_capa:
        linha_capa = {
            "@fotofundo": rf"C:\Users\Administrador\Documents\fotosref\{referencia_capa_val}.jpg" if referencia_capa_val else "",
            "@logo": rf"C:\Users\Administrador\Documents\Sistemas\PDFgenerator\static\logos\{logo_escolhida}.png" if want_logo and logo_escolhida else "",
            "@sublogo": rf"C:\Users\Administrador\Documents\Sistemas\PDFgenerator\static\logos\{sublogo_escolhida}.png" if want_sublogo and sublogo_escolhida else ""
        }
        df_capa = pd.DataFrame([linha_capa])
        df_capa.to_csv(CSV_CAPA_PATH, index=False, sep=",", encoding="utf-16")
    else:
        # garante que arquivo anterior não atrapalhe
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

    # escolher qual JSX usar
    script_to_run = escolher_script(want_capa, want_contracapa)

    # executar InDesign chamando o JSX adequado
    sucesso = executar_indesign_with_jsx(script_to_run)

    if sucesso:
        time.sleep(5)
        return redirect(url_for('visualizar'))
    else:
        return jsonify({"erro": "Falha ao executar o InDesign (VBS/JSX)."}), 500
    


@app.route("/download")
@login_required
def download_pdf():
    if not os.path.exists(PDF_PATH):
        return jsonify({"erro": "PDF não encontrado."}), 404

    nome = session.get("nome_arquivo_escolhido", "arquivo_final")
    # sanitize: opcionalmente remova espaços/char inválidos do nome
    nome = "".join(c for c in nome if c.isalnum() or c in (" ", "-", "_")).strip()
    if not nome:
        nome = "arquivo_final"

    return send_file(
        PDF_PATH,
        as_attachment=True,
        download_name=f"{nome}.pdf"
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)