import os
from datetime import datetime
import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, jsonify

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MODELO_EXCEL = os.path.join(BASE_DIR, "ControleOcorrencias.xlsx")
DADOS_EXCEL = os.path.join(BASE_DIR, "dados_ocorrencias.xlsx")

app = Flask(__name__)

# -------------------- Funções auxiliares --------------------
def ler_professores():
    try:
        return pd.read_excel(MODELO_EXCEL, sheet_name="Professores")["Professor"].dropna().tolist()
    except Exception:
        return []

def ler_salas():
    try:
        return pd.read_excel(MODELO_EXCEL, sheet_name="Salas")["Sala"].dropna().tolist()
    except Exception:
        return []

def ler_alunos(sala=None):
    try:
        df = pd.read_excel(MODELO_EXCEL, sheet_name="Alunos")
        return df[df["Sala"] == sala]["Aluno"].dropna().tolist() if sala else df["Aluno"].dropna().tolist()
    except Exception:
        return []

def ler_tutor(aluno):
    try:
        df = pd.read_excel(MODELO_EXCEL, sheet_name="Alunos")
        linha = df[df["Aluno"] == aluno]
        if not linha.empty:
            return linha.iloc[0].get("Tutor", "")
    except Exception:
        pass
    return ""

def carregar_dados():
    cols = [
        "Nº Ocorrência", "Data Criação", "Hora Criação",
        "Professor", "Sala", "Aluno", "Tutor",
        "Descrição da Ocorrência",
        "Atendimento Professor", "Atendimento Tutor",
        "Atendimento Coordenação", "Atendimento Gestão",
        "FlagTutor", "FlagCoord", "FlagGestao",
        "Data Atendimento Tutor", "Data Atendimento Coord", "Data Atendimento Gestao",
        "Status"
    ]
    if not os.path.exists(DADOS_EXCEL):
        return pd.DataFrame(columns=cols)
    try:
        df = pd.read_excel(DADOS_EXCEL)
        for c in cols:
            if c not in df.columns:
                df[c] = ""
        return df[cols]
    except Exception:
        return pd.DataFrame(columns=cols)

def salvar_dados(df):
    df.to_excel(DADOS_EXCEL, index=False)

def proximo_numero():
    df = carregar_dados()
    if df.empty:
        return 1
    try:
        return int(df["Nº Ocorrência"].max()) + 1
    except Exception:
        return len(df) + 1

def calc_status(ft, fc, fg):
    if ft == "Não" and fc == "Não" and fg == "Não":
        return "Finalizada"
    return "Em Atendimento"

# -------------------- Rotas --------------------
@app.route("/")
def home():
    return render_template("home.html")

@app.route("/ocorrencias")
def index():
    tutor = request.args.get("tutor", "").strip()
    status = request.args.get("status", "").strip()
    sala = request.args.get("sala", "").strip()
    aluno = request.args.get("aluno", "").strip()

    df = carregar_dados()
    if not df.empty:
        df["Status"] = df.apply(lambda r: calc_status(r["FlagTutor"], r["FlagCoord"], r["FlagGestao"]), axis=1)

        if tutor:
            df = df[df["Tutor"] == tutor]
        if status:
            df = df[df["Status"] == status]
        if sala:
            df = df[df["Sala"] == sala]
        if aluno:
            df = df[df["Aluno"] == aluno]

    try:
        tutores_ref = pd.read_excel(MODELO_EXCEL, sheet_name="Alunos")["Tutor"].dropna().unique().tolist()
    except Exception:
        tutores_ref = []

    registros = [] if df.empty else df.to_dict(orient="records")
    return render_template("index.html", registros=registros, tutores=tutores_ref, salas=ler_salas())

@app.route("/nova")
def nova():
    return render_template("nova.html", professores=ler_professores(), salas=ler_salas())

@app.route("/salvar", methods=["POST"])
def salvar():
    df = carregar_dados()
    numero = proximo_numero()
    agora = datetime.now()

    nova = {
        "Nº Ocorrência": numero,
        "Data Criação": agora.strftime("%Y-%m-%d"),
        "Hora Criação": agora.strftime("%H:%M:%S"),
        "Professor": request.form.get("professor", ""),
        "Sala": request.form.get("sala", ""),
        "Aluno": request.form.get("aluno", ""),
        "Tutor": request.form.get("tutor", ""),
        "Descrição da Ocorrência": request.form.get("descricao", ""),
        "Atendimento Professor": request.form.get("at_professor", ""),
        "Atendimento Tutor": "",
        "Atendimento Coordenação": "",
        "Atendimento Gestão": "",
        "FlagTutor": "Não",
        "FlagCoord": "Não",
        "FlagGestao": "Não",
        "Data Atendimento Tutor": "",
        "Data Atendimento Coord": "",
        "Data Atendimento Gestao": "",
        "Status": "Em Atendimento",
    }
    df = pd.concat([df, pd.DataFrame([nova])], ignore_index=True)
    salvar_dados(df)
    return redirect(url_for("index"))

@app.route("/abrir_pendencia/<int:oid>/<papel>")
def abrir_pendencia(oid, papel):
    df = carregar_dados()
    if oid not in df["Nº Ocorrência"].values:
        return "Ocorrência não encontrada", 404
    if papel == "tutor":
        df.loc[df["Nº Ocorrência"] == oid, "FlagTutor"] = "Sim"
    elif papel == "coord":
        df.loc[df["Nº Ocorrência"] == oid, "FlagCoord"] = "Sim"
    elif papel == "gestao":
        df.loc[df["Nº Ocorrência"] == oid, "FlagGestao"] = "Sim"

    linha = df[df["Nº Ocorrência"] == oid].iloc[0]
    df.loc[df["Nº Ocorrência"] == oid, "Status"] = calc_status(
        linha["FlagTutor"], linha["FlagCoord"], linha["FlagGestao"]
    )
    salvar_dados(df)
    return redirect(url_for("editar", oid=oid, campo=papel))

@app.route("/editar/<int:oid>", methods=["GET", "POST"])
def editar(oid):
    campo = request.args.get("campo", "").strip()
    df = carregar_dados()
    linha = df[df["Nº Ocorrência"] == oid]
    if linha.empty:
        return "Ocorrência não encontrada", 404

    registro = linha.iloc[0].to_dict()

    if request.method == "POST":
        if campo in ["edit", "tutor", "coord", "gestao"]:
            if campo in ["edit", "tutor"]:
                df.loc[df["Nº Ocorrência"] == oid, "Atendimento Tutor"] = request.form.get("at_tutor", "")
                if campo == "tutor":
                    df.loc[df["Nº Ocorrência"] == oid, "FlagTutor"] = "Não"
            if campo in ["edit", "coord"]:
                df.loc[df["Nº Ocorrência"] == oid, "Atendimento Coordenação"] = request.form.get("at_coord", "")
                if campo == "coord":
                    df.loc[df["Nº Ocorrência"] == oid, "FlagCoord"] = "Não"
            if campo in ["edit", "gestao"]:
                df.loc[df["Nº Ocorrência"] == oid, "Atendimento Gestão"] = request.form.get("at_gestao", "")
                if campo == "gestao":
                    df.loc[df["Nº Ocorrência"] == oid, "FlagGestao"] = "Não"
            if campo == "edit":
                df.loc[df["Nº Ocorrência"] == oid, "Atendimento Professor"] = request.form.get("at_professor", "")

        linha_atual = df[df["Nº Ocorrência"] == oid].iloc[0]
        df.loc[df["Nº Ocorrência"] == oid, "Status"] = calc_status(
            linha_atual["FlagTutor"], linha_atual["FlagCoord"], linha_atual["FlagGestao"]
        )
        salvar_dados(df)
        return redirect(url_for("index"))

    permissoes = {
        "professor": campo == "edit",
        "tutor": campo in ["edit", "tutor"],
        "coord": campo in ["edit", "coord"],
        "gestao": campo in ["edit", "gestao"]
    }
    return render_template("editar.html", ocorrencia=registro, permissoes=permissoes)

@app.route("/api/alunos/<sala>")
def api_alunos(sala):
    return jsonify(ler_alunos(sala))

@app.route("/api/tutor/<aluno>")
def api_tutor(aluno):
    return jsonify({"tutor": ler_tutor(aluno)})

@app.route("/tutoria")
def tutoria():
    return "<h2 style='text-align:center;margin-top:40px'>Página de Tutoria em desenvolvimento...</h2>"

@app.route("/relatorio")
def relatorio():
    # Exemplo simples: depois você coloca o código real para gerar o relatório
    return "<h3 style='text-align:center'>Página de Relatório</h3>"

if __name__ == "__main__":
    app.run(debug=True)
