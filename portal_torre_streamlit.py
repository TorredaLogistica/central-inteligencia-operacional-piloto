import hashlib
import hmac
import json
import os
import secrets
import sqlite3
from urllib.parse import quote
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
from pathlib import Path

import streamlit as st
import pandas as pd

DB_PATH = Path(os.getenv("TORRE_DB_PATH", "torre_usuarios.db"))
USUARIOS_JSON = Path(os.getenv("TORRE_USUARIOS_JSON", "usuarios.json"))
ITERACOES = 600_000
FUSO_BRASILIA = ZoneInfo("America/Sao_Paulo")

# Catálogo dos indicadores. Para disponibilizar um indicador, preencha a URL.
# Os títulos foram mantidos alinhados à Central HTML atual.
INDICADORES = {
    "Armazenagem": [
        {"titulo": "Chatbot Torre Logística", "url": "https://chatbot-logistica-e4cpfbcta3qyqchopsdjeg.streamlit.app/", "icone": "🤖"},
        {"titulo": "Separação e Faturamento", "url": "https://dashboard-slaseparacaofaturamento-mdnfzinkaebwzysne83ewp.streamlit.app/", "icone": "📊"},
        {"titulo": "Pedidos para LPs", "url": "https://pedidoslpsaas-y44bkbmcg4kon8fogbro34.streamlit.app/", "icone": "📦"},
        {"titulo": "Resultado do DRE", "url": "https://resultadodre-lk6rh4ahefeuwfhwrg2ioc.streamlit.app/", "icone": "💰"},
        {"titulo": "Valores dos EAs", "url": "https://valorestoques-eas-73bxfsks3rnoxjo44fuqm7.streamlit.app/", "icone": "📈"},
        {"titulo": "Atendimento de OVs nos TLs", "url": "https://atendimento-de-ovs-nos-tls-in3rykeacnvjxedhb7r9zc.streamlit.app/", "icone": "📋"},
        {"titulo": "Taxa de Ocupação dos CDs", "url": "https://taxadeocupacaodoscds-tfx8ftu78n46vhvn5dxc7k.streamlit.app/", "icone": "🏭"},
        {"titulo": "Pedidos Canal Vermelho", "url": "CANAL_VERMELHO", "icone": "⚡"},
        {"titulo": "Pedidos LPs e AAs NFs não Agrupadas", "url": "https://nfs-nao-agrupadas-juqvjn8nhbzuhzzdfwknbl.streamlit.app/", "icone": "🧾"},
        {"titulo": "Faturas dos OPLs", "url": "https://faturasdosopls-fdgzskwvbciwkgubekfjcz.streamlit.app/", "icone": "💵"},
        {"titulo": "Recebimento de Usados", "url": "https://controlederecebimentodeusados-ucfnrvqwuceiet5q7tt4wn.streamlit.app/", "icone": "♻️"},
        {"titulo": "Nível de Serviços dos OPLs", "url": "https://niveldeservicoopls-sgrryyugyheukmxp2xtzp8.streamlit.app/", "icone": "🚚"},
        {"titulo": "Forecast e Realizado", "url": "https://forecasterealizado-kebvdtavq5yc8s9kwrfqiu.streamlit.app/", "icone": "🔢"},
        {"titulo": "Simulação de Pedidos", "url": "https://simulacaopedidos-myjtjrm3nklxzprutjpbd5.streamlit.app/", "icone": "🧮"},
    ],
    "Triagem": [],
    "Reversa": [],
}

CANAL_VERMELHO_URL = "https://cuencjwy3ahhnzymzsspuc.streamlit.app/"
CANAL_VERMELHO_APP_ID = "canal_vermelho"
CANAL_VERMELHO_SHARED_KEY = "b9342075f69fbf07834993550e178cb29eec8346a37259c03c8444e7df541e01"

st.set_page_config(page_title="Claro | Central de Inteligência Operacional", page_icon="🔴", layout="wide")


def agora_iso():
    # Armazena o instante em UTC para manter ordenação e auditoria consistentes.
    return datetime.now(timezone.utc).isoformat()


def data_hora_brasilia(valor):
    if not valor:
        return "-"
    try:
        dt = datetime.fromisoformat(str(valor).replace("Z", "+00:00"))
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        return dt.astimezone(FUSO_BRASILIA).strftime("%d/%m/%Y %H:%M:%S")
    except (TypeError, ValueError):
        return str(valor)


def normalizar_usuario(valor: str) -> str:
    import unicodedata
    valor = unicodedata.normalize("NFD", str(valor or "").strip())
    valor = "".join(c for c in valor if unicodedata.category(c) != "Mn")
    return "".join(c for c in valor.upper() if ("A" <= c <= "Z") or ("0" <= c <= "9"))


def normalizar_email(valor: str) -> str:
    return (valor or "").strip().lower()


def sha256(valor: str) -> str:
    return hashlib.sha256(valor.encode("utf-8")).hexdigest()


def gerar_hash_senha(senha: str, salt_hex: str | None = None):
    salt_hex = salt_hex or secrets.token_hex(16)
    derivado = hashlib.pbkdf2_hmac("sha256", senha.encode(), bytes.fromhex(salt_hex), ITERACOES, dklen=32)
    return derivado.hex(), salt_hex


def conferir_senha(senha: str, salt_hex: str, hash_esperado: str, iteracoes: int):
    derivado = hashlib.pbkdf2_hmac("sha256", senha.encode(), bytes.fromhex(salt_hex), int(iteracoes), dklen=32).hex()
    return hmac.compare_digest(derivado, hash_esperado)


def conectar():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def gerar_url_canal_vermelho():
    ts = str(int(datetime.now(timezone.utc).timestamp()))
    nonce = secrets.token_hex(16)
    mensagem = f"{CANAL_VERMELHO_APP_ID}|{ts}|{nonce}"
    assinatura = hmac.new(
        CANAL_VERMELHO_SHARED_KEY.encode("utf-8"),
        mensagem.encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()
    return (
        f"{CANAL_VERMELHO_URL}?portal_ts={quote(ts)}"
        f"&portal_nonce={quote(nonce)}&portal_sig={quote(assinatura)}"
    )


def exibir_central_indicadores():
    usuario_atual = st.session_state.usuario_logado or {}
    nome = usuario_atual.get("nome_completo") or usuario_atual.get("usuario") or "Usuário"

    usuario_exibicao = usuario_atual.get("usuario") or nome
    esquerda_controles, centro_controles, direita_controles = st.columns([2.7, 1.1, 1.2])
    with esquerda_controles:
        st.markdown('<span class="header-area-anchor menu-model-anchor"></span>', unsafe_allow_html=True)
        area = st.segmented_control(
            "Área operacional",
            list(INDICADORES.keys()),
            default="Armazenagem",
            key="area_indicadores_padrao",
            label_visibility="collapsed",
        ) or "Armazenagem"
    with direita_controles:
        st.markdown(f'<div class="header-user-name">Usuário: {usuario_exibicao}</div><span class="exit-model-anchor"></span>', unsafe_allow_html=True)
        if st.button("Sair", use_container_width=True, type="secondary", key="btn_sair_header"):
            st.session_state.usuario_logado = None
            st.rerun()

    indicadores = INDICADORES.get(area, [])

    if not indicadores:
        st.info(f"Nenhum indicador disponível para a área de {area} no momento.")
        return

    for inicio in range(0, len(indicadores), 5):
        linha = indicadores[inicio:inicio + 5]
        colunas = st.columns(5, gap="medium")
        for coluna, indicador in zip(colunas, linha):
            with coluna:
                destino = gerar_url_canal_vermelho() if indicador["url"] == "CANAL_VERMELHO" else indicador["url"]
                titulo = indicador["titulo"]
                icone = indicador["icone"]
                card_html = f"""<a class="indicador-card-link" href="{destino}" target="_blank" rel="noopener noreferrer">
                    <span class="indicador-card-acento"></span>
                    <span class="indicador-card-simbolo" aria-hidden="true">{icone}</span>
                    <span class="indicador-card-titulo">{titulo}</span>
                </a>"""
                st.markdown(card_html, unsafe_allow_html=True)


def iniciar_banco():
    with conectar() as con:
        con.executescript("""
        CREATE TABLE IF NOT EXISTS usuarios (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome_completo TEXT,
            usuario TEXT,
            email TEXT,
            nome_hash TEXT UNIQUE NOT NULL,
            email_hash TEXT UNIQUE NOT NULL,
            senha_hash TEXT NOT NULL,
            salt TEXT NOT NULL,
            iteracoes INTEGER NOT NULL DEFAULT 600000,
            codigo_dispositivo_hash TEXT,
            ativo INTEGER NOT NULL DEFAULT 0,
            perfil TEXT NOT NULL DEFAULT 'USUARIO',
            criado_em TEXT NOT NULL,
            aprovado_em TEXT,
            atualizado_em TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS solicitacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tipo TEXT NOT NULL,
            usuario_id INTEGER,
            nome_completo TEXT,
            usuario TEXT NOT NULL,
            email TEXT NOT NULL,
            nome_hash TEXT NOT NULL,
            email_hash TEXT NOT NULL,
            senha_hash_pendente TEXT NOT NULL,
            salt_pendente TEXT NOT NULL,
            iteracoes INTEGER NOT NULL DEFAULT 600000,
            codigo_dispositivo_hash TEXT,
            status TEXT NOT NULL DEFAULT 'PENDENTE',
            criado_em TEXT NOT NULL,
            decidido_em TEXT,
            decidido_por TEXT,
            observacao TEXT,
            FOREIGN KEY(usuario_id) REFERENCES usuarios(id)
        );
        CREATE TABLE IF NOT EXISTS auditoria (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            acao TEXT NOT NULL,
            referencia TEXT,
            administrador TEXT,
            criado_em TEXT NOT NULL
        );
        """)


def importar_json_legado():
    if not USUARIOS_JSON.exists():
        return
    try:
        dados = json.loads(USUARIOS_JSON.read_text(encoding="utf-8"))
    except Exception:
        return
    with conectar() as con:
        quantidade = con.execute("SELECT COUNT(*) FROM usuarios").fetchone()[0]
        if quantidade:
            return
        for u in dados.get("usuarios", []):
            con.execute("""INSERT OR IGNORE INTO usuarios
                (nome_hash,email_hash,senha_hash,salt,iteracoes,codigo_dispositivo_hash,ativo,perfil,criado_em,aprovado_em,atualizado_em)
                VALUES(?,?,?,?,?,?,?,?,?,?,?)""", (
                u["nome_hash"], u["email_hash"], u["senha_hash"], u["salt"], u.get("iteracoes", ITERACOES),
                u.get("codigo_dispositivo_hash"), 1 if u.get("ativo") else 0, "USUARIO",
                u.get("aprovado_em") or agora_iso(), u.get("aprovado_em"), agora_iso()
            ))


def criar_solicitacao_cadastro(nome, usuario, email, senha):
    usuario = normalizar_usuario(usuario)
    email = normalizar_email(email)
    if not usuario or not email.endswith("@claro.com.br") or len(senha) < 8:
        raise ValueError("Revise o usuário, o e-mail corporativo e a senha.")
    nome_hash, email_hash = sha256(usuario.lower()), sha256(email)
    senha_hash, salt = gerar_hash_senha(senha)
    token = "-".join([secrets.token_hex(2).upper() for _ in range(4)])
    token_hash = sha256(token.replace("-", ""))
    with conectar() as con:
        existe = con.execute("SELECT 1 FROM usuarios WHERE nome_hash=? OR email_hash=?", (nome_hash, email_hash)).fetchone()
        pendente = con.execute("SELECT 1 FROM solicitacoes WHERE status='PENDENTE' AND (nome_hash=? OR email_hash=?)", (nome_hash, email_hash)).fetchone()
        if existe or pendente:
            raise ValueError("Já existe cadastro ou solicitação pendente para os dados informados.")
        con.execute("""INSERT INTO solicitacoes
            (tipo,nome_completo,usuario,email,nome_hash,email_hash,senha_hash_pendente,salt_pendente,iteracoes,codigo_dispositivo_hash,status,criado_em)
            VALUES('CADASTRO',?,?,?,?,?,?,?,?,?,'PENDENTE',?)""",
            (nome.strip(), usuario, email, nome_hash, email_hash, senha_hash, salt, ITERACOES, token_hash, agora_iso()))
    return token


def criar_solicitacao_senha(usuario, email, nova_senha):
    usuario = normalizar_usuario(usuario)
    email = normalizar_email(email)
    if len(nova_senha) < 8:
        raise ValueError("A nova senha deve ter no mínimo 8 caracteres.")
    nome_hash, email_hash = sha256(usuario.lower()), sha256(email)
    with conectar() as con:
        u = con.execute("SELECT * FROM usuarios WHERE nome_hash=? AND email_hash=? AND ativo=1", (nome_hash, email_hash)).fetchone()
        # Mensagem externa deve continuar genérica; ausência não revela cadastro.
        if not u:
            return
        pendente = con.execute("SELECT 1 FROM solicitacoes WHERE status='PENDENTE' AND tipo='REDEFINICAO_SENHA' AND usuario_id=?", (u["id"],)).fetchone()
        if pendente:
            return
        senha_hash, salt = gerar_hash_senha(nova_senha)
        con.execute("""INSERT INTO solicitacoes
            (tipo,usuario_id,nome_completo,usuario,email,nome_hash,email_hash,senha_hash_pendente,salt_pendente,iteracoes,status,criado_em)
            VALUES('REDEFINICAO_SENHA',?,?,?,?,?,?,?,?,?,'PENDENTE',?)""",
            (u["id"], u["nome_completo"], usuario, email, nome_hash, email_hash, senha_hash, salt, ITERACOES, agora_iso()))


def autenticar_usuario(usuario, email, senha):
    nome_hash = sha256(normalizar_usuario(usuario).lower())
    email_hash = sha256(normalizar_email(email))
    with conectar() as con:
        u = con.execute("SELECT * FROM usuarios WHERE nome_hash=? AND email_hash=? AND ativo=1", (nome_hash, email_hash)).fetchone()
    if not u or not conferir_senha(senha, u["salt"], u["senha_hash"], u["iteracoes"]):
        return None
    return dict(u)


def credenciais_admin_validas(usuario, senha):
    admin_user = st.secrets.get("ADMIN_USUARIO", "")
    admin_hash = st.secrets.get("ADMIN_SENHA_HASH", "")
    return bool(admin_user and admin_hash and hmac.compare_digest(usuario, admin_user) and hmac.compare_digest(sha256(senha), admin_hash))


def decidir_solicitacao(solicitacao_id, aprovar, administrador, observacao=""):
    with conectar() as con:
        s = con.execute("SELECT * FROM solicitacoes WHERE id=? AND status='PENDENTE'", (solicitacao_id,)).fetchone()
        if not s:
            return
        status = "APROVADO" if aprovar else "REJEITADO"
        if aprovar and s["tipo"] == "CADASTRO":
            con.execute("""INSERT INTO usuarios
                (nome_completo,usuario,email,nome_hash,email_hash,senha_hash,salt,iteracoes,codigo_dispositivo_hash,ativo,perfil,criado_em,aprovado_em,atualizado_em)
                VALUES(?,?,?,?,?,?,?,?,?,1,'USUARIO',?,?,?)""",
                (s["nome_completo"],s["usuario"],s["email"],s["nome_hash"],s["email_hash"],s["senha_hash_pendente"],s["salt_pendente"],s["iteracoes"],s["codigo_dispositivo_hash"],s["criado_em"],agora_iso(),agora_iso()))
        elif aprovar and s["tipo"] == "REDEFINICAO_SENHA":
            con.execute("UPDATE usuarios SET senha_hash=?,salt=?,iteracoes=?,atualizado_em=? WHERE id=?",
                        (s["senha_hash_pendente"],s["salt_pendente"],s["iteracoes"],agora_iso(),s["usuario_id"]))
        con.execute("UPDATE solicitacoes SET status=?,decidido_em=?,decidido_por=?,observacao=? WHERE id=?",
                    (status,agora_iso(),administrador,observacao,solicitacao_id))
        con.execute("INSERT INTO auditoria(acao,referencia,administrador,criado_em) VALUES(?,?,?,?)",
                    (f"{status}_{s['tipo']}",str(solicitacao_id),administrador,agora_iso()))


def listar_usuarios():
    with conectar() as con:
        return [dict(x) for x in con.execute(
            """SELECT id,nome_completo,usuario,email,ativo,perfil,criado_em,aprovado_em,atualizado_em
               FROM usuarios ORDER BY COALESCE(nome_completo,usuario,email,nome_hash)"""
        ).fetchall()]


def editar_cadastro_usuario(usuario_id, nome_completo, usuario, email, administrador):
    nome_completo = str(nome_completo or "").strip()
    usuario = normalizar_usuario(usuario)
    email = normalizar_email(email)
    if not nome_completo:
        raise ValueError("Informe o nome completo.")
    if not usuario:
        raise ValueError("Informe um nome de usuário válido.")
    if not email.endswith("@claro.com.br"):
        raise ValueError("Informe um e-mail corporativo válido.")

    novo_nome_hash = sha256(usuario.lower())
    novo_email_hash = sha256(email)
    with conectar() as con:
        atual = con.execute(
            "SELECT id,nome_completo,usuario,email,nome_hash,email_hash FROM usuarios WHERE id=?",
            (usuario_id,),
        ).fetchone()
        if not atual:
            raise ValueError("Cadastro não localizado.")

        duplicado = con.execute(
            """SELECT id FROM usuarios
               WHERE id<>? AND (nome_hash=? OR email_hash=?)""",
            (usuario_id, novo_nome_hash, novo_email_hash),
        ).fetchone()
        if duplicado:
            raise ValueError("O usuário ou o e-mail informado já pertence a outro cadastro.")

        referencia = json.dumps({
            "usuario_id": usuario_id,
            "nome_anterior": atual["nome_completo"],
            "usuario_anterior": atual["usuario"],
            "email_anterior": atual["email"],
            "nome_novo": nome_completo,
            "usuario_novo": usuario,
            "email_novo": email,
        }, ensure_ascii=False)
        con.execute(
            """UPDATE usuarios
               SET nome_completo=?,usuario=?,email=?,nome_hash=?,email_hash=?,atualizado_em=?
               WHERE id=?""",
            (nome_completo, usuario, email, novo_nome_hash, novo_email_hash, agora_iso(), usuario_id),
        )
        con.execute(
            "INSERT INTO auditoria(acao,referencia,administrador,criado_em) VALUES(?,?,?,?)",
            ("EDITAR_CADASTRO_USUARIO", referencia, administrador, agora_iso()),
        )
    return True


def alterar_status_usuario(usuario_id, ativo, administrador):
    with conectar() as con:
        usuario = con.execute("SELECT id,usuario,email,ativo FROM usuarios WHERE id=?", (usuario_id,)).fetchone()
        if not usuario:
            return False
        con.execute("UPDATE usuarios SET ativo=?,atualizado_em=? WHERE id=?", (1 if ativo else 0, agora_iso(), usuario_id))
        con.execute("INSERT INTO auditoria(acao,referencia,administrador,criado_em) VALUES(?,?,?,?)",
                    ("REATIVAR_USUARIO" if ativo else "DESATIVAR_USUARIO", str(usuario_id), administrador, agora_iso()))
    return True


def excluir_usuario(usuario_id, administrador):
    with conectar() as con:
        usuario = con.execute("SELECT id FROM usuarios WHERE id=?", (usuario_id,)).fetchone()
        if not usuario:
            return False
        # Preserva o histórico das solicitações, removendo apenas o vínculo técnico.
        con.execute("UPDATE solicitacoes SET usuario_id=NULL WHERE usuario_id=?", (usuario_id,))
        con.execute("DELETE FROM usuarios WHERE id=?", (usuario_id,))
        con.execute("INSERT INTO auditoria(acao,referencia,administrador,criado_em) VALUES(?,?,?,?)",
                    ("EXCLUIR_USUARIO", str(usuario_id), administrador, agora_iso()))
    return True


def importar_usuarios_json_manual(arquivo, administrador):
    try:
        dados = json.loads(arquivo.getvalue().decode("utf-8-sig"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError("O arquivo selecionado não contém um JSON válido.") from exc

    usuarios = dados.get("usuarios") if isinstance(dados, dict) else None
    if not isinstance(usuarios, list):
        raise ValueError("O JSON precisa conter a chave 'usuarios' no formato de lista.")

    obrigatorios = {"nome_hash", "email_hash", "senha_hash", "salt"}
    analisados = len(usuarios)
    incluidos = 0
    atualizados = 0
    ignorados = 0
    erros = []

    with conectar() as con:
        for indice, item in enumerate(usuarios, start=1):
            if not isinstance(item, dict):
                ignorados += 1
                erros.append(f"Registro {indice}: estrutura inválida.")
                continue
            ausentes = sorted(campo for campo in obrigatorios if not item.get(campo))
            if ausentes:
                ignorados += 1
                erros.append(f"Registro {indice}: campos ausentes: {', '.join(ausentes)}.")
                continue
            try:
                int(item.get("iteracoes", ITERACOES))
                bytes.fromhex(str(item["salt"]))
            except (TypeError, ValueError):
                ignorados += 1
                erros.append(f"Registro {indice}: salt ou iterações inválidos.")
                continue

            existente = con.execute(
                "SELECT id FROM usuarios WHERE nome_hash=? OR email_hash=?",
                (item["nome_hash"], item["email_hash"]),
            ).fetchone()
            agora = agora_iso()
            aprovado_em = item.get("aprovado_em") or agora
            valores = (
                item["nome_hash"],
                item["email_hash"],
                item["senha_hash"],
                item["salt"],
                int(item.get("iteracoes", ITERACOES)),
                item.get("codigo_dispositivo_hash"),
                1 if item.get("ativo", True) else 0,
                aprovado_em,
                agora,
            )
            if existente:
                con.execute(
                    """UPDATE usuarios
                       SET nome_hash=?,email_hash=?,senha_hash=?,salt=?,iteracoes=?,
                           codigo_dispositivo_hash=?,ativo=?,aprovado_em=?,atualizado_em=?
                       WHERE id=?""",
                    valores + (existente["id"],),
                )
                atualizados += 1
            else:
                con.execute(
                    """INSERT INTO usuarios
                       (nome_hash,email_hash,senha_hash,salt,iteracoes,codigo_dispositivo_hash,
                        ativo,perfil,criado_em,aprovado_em,atualizado_em)
                       VALUES(?,?,?,?,?,?,?,'USUARIO',?,?,?)""",
                    valores[:7] + (aprovado_em, aprovado_em, agora),
                )
                incluidos += 1

        con.execute(
            "INSERT INTO auditoria(acao,referencia,administrador,criado_em) VALUES(?,?,?,?)",
            (
                "IMPORTAR_USUARIOS_JSON",
                f"analisados={analisados};incluidos={incluidos};atualizados={atualizados};ignorados={ignorados}",
                administrador,
                agora_iso(),
            ),
        )

    return {
        "analisados": analisados,
        "incluidos": incluidos,
        "atualizados": atualizados,
        "ignorados": ignorados,
        "erros": erros,
    }


def complementar_cadastros_xlsx(arquivo, administrador):
    try:
        planilha = pd.read_excel(arquivo, sheet_name="Usuarios", engine="openpyxl", dtype=str)
    except ValueError as exc:
        raise ValueError("A planilha precisa conter uma guia chamada 'Usuarios'.") from exc
    except Exception as exc:
        raise ValueError("Não foi possível ler o arquivo XLSX selecionado.") from exc

    planilha.columns = [str(coluna).strip().lower() for coluna in planilha.columns]
    obrigatorias = ["nome_completo", "usuario", "email"]
    ausentes = [coluna for coluna in obrigatorias if coluna not in planilha.columns]
    if ausentes:
        raise ValueError("Colunas obrigatórias ausentes: " + ", ".join(ausentes) + ".")

    planilha = planilha[obrigatorias].fillna("")
    analisados = len(planilha)
    complementados = ja_completos = nao_localizados = ignorados = 0
    erros = []
    with conectar() as con:
        for numero_linha, linha in enumerate(planilha.itertuples(index=False), start=2):
            nome_completo = str(linha.nome_completo).strip()
            usuario = normalizar_usuario(str(linha.usuario))
            email = normalizar_email(str(linha.email))
            if not nome_completo or not usuario or not email:
                ignorados += 1
                erros.append(f"Linha {numero_linha}: nome, usuário ou e-mail não preenchido.")
                continue
            if not email.endswith("@claro.com.br"):
                ignorados += 1
                erros.append(f"Linha {numero_linha}: e-mail corporativo inválido.")
                continue
            cadastro = con.execute(
                "SELECT id,nome_completo,usuario,email FROM usuarios WHERE nome_hash=? AND email_hash=?",
                (sha256(usuario.lower()), sha256(email)),
            ).fetchone()
            if not cadastro:
                nao_localizados += 1
                erros.append(f"Linha {numero_linha}: cadastro legado não localizado pelos dados informados.")
                continue
            valores_atuais = (
                str(cadastro["nome_completo"] or ""),
                str(cadastro["usuario"] or ""),
                str(cadastro["email"] or ""),
            )
            valores_planilha = (nome_completo, usuario, email)
            if valores_atuais == valores_planilha:
                ja_completos += 1
                continue

            con.execute(
                "UPDATE usuarios SET nome_completo=?,usuario=?,email=?,atualizado_em=? WHERE id=?",
                (nome_completo, usuario, email, agora_iso(), cadastro["id"]),
            )
            complementados += 1
        con.execute(
            "INSERT INTO auditoria(acao,referencia,administrador,criado_em) VALUES(?,?,?,?)",
            ("COMPLEMENTAR_CADASTROS_XLSX",
             f"analisados={analisados};complementados={complementados};ja_completos={ja_completos};nao_localizados={nao_localizados};ignorados={ignorados}",
             administrador, agora_iso()),
        )
    return {"analisados": analisados, "complementados": complementados,
            "ja_completos": ja_completos, "nao_localizados": nao_localizados,
            "ignorados": ignorados, "erros": erros}


def exportar_json_compatibilidade():
    with conectar() as con:
        linhas = con.execute("SELECT nome_hash,email_hash,senha_hash,salt,iteracoes,codigo_dispositivo_hash,ativo,aprovado_em FROM usuarios").fetchall()
    return json.dumps({"versao": 28, "usuarios": [dict(x) for x in linhas], "atualizado_em": agora_iso()}, ensure_ascii=False, indent=2)


iniciar_banco()
importar_json_legado()

st.markdown("""
<style>
.stApp{background:#eef0f3}
.block-container{max-width:1900px;padding-top:2.35rem;padding-left:1.25rem;padding-right:1.25rem}
.portal-head{position:relative;border-radius:0 0 24px 24px;overflow:hidden;background:linear-gradient(180deg,#b51f25 0%,#cf2b25 48%,#f47b45 100%);box-shadow:0 12px 30px rgba(103,0,0,.17);padding:30px 26px 18px;margin:-1.35rem -1.25rem 1.25rem;color:#fff}
.portal-head:after{content:"";position:absolute;width:420px;height:220px;right:-130px;top:-120px;border-radius:50%;background:radial-gradient(circle,rgba(255,255,255,.14),transparent 68%)}
.portal-brand{position:relative;z-index:1;display:flex;align-items:center;justify-content:center;gap:20px;min-height:82px}
.portal-logo-claro{width:88px;height:88px;object-fit:contain;flex:0 0 auto}
.portal-logo-logistica{width:145px;height:70px;object-fit:contain;flex:0 0 auto}
.portal-sep{width:1px;height:56px;background:rgba(255,255,255,.58);flex:0 0 1px}
.portal-title{font-size:30px;font-weight:850;line-height:1.12;letter-spacing:-.6px;text-align:center;white-space:nowrap}
.portal-subtitle{position:relative;z-index:1;text-align:center;margin-top:2px;color:rgba(255,255,255,.86);font-size:13px;font-weight:600}
.welcome-strip{padding:15px 18px;border-radius:16px;background:linear-gradient(90deg,#fff,#fff7f4);border:1px solid rgba(218,41,28,.14);box-shadow:0 8px 22px rgba(70,20,20,.08);color:#333;font-size:1rem;font-weight:750}
.pendencia{padding:.8rem 1rem;border-radius:12px;background:#fff3cd;border:1px solid #ffec99;color:#7a5200;font-weight:700}
/* Modelo 1: cards modernos com ícones */
.indicador-card-link{position:relative;display:flex;min-height:182px;padding:26px 18px 22px;border:1px solid rgba(218,41,28,.14);border-radius:22px;overflow:hidden;background:linear-gradient(145deg,rgba(255,255,255,.99),rgba(255,246,243,.97));box-shadow:0 10px 28px rgba(79,20,20,.10);text-decoration:none!important;color:#292929!important;flex-direction:column;align-items:center;justify-content:center;transition:transform .22s ease,box-shadow .22s ease,border-color .22s ease}
.indicador-card-link:before{content:"";position:absolute;width:140px;height:140px;right:-66px;top:-66px;border-radius:50%;background:radial-gradient(circle,rgba(244,123,69,.27),rgba(218,41,28,.07) 55%,transparent 71%)}
.indicador-card-link:hover{transform:translateY(-6px);box-shadow:0 18px 38px rgba(122,25,25,.18);border-color:rgba(218,41,28,.36)}
.indicador-card-acento{position:absolute;left:0;top:0;width:100%;height:5px;background:linear-gradient(90deg,#b51f25,#da291c 50%,#f47b45)}
.indicador-card-simbolo{display:flex;align-items:center;justify-content:center;width:64px;height:64px;border-radius:20px;background:linear-gradient(145deg,#ed3025,#b51f25);color:#fff;font-size:30px;line-height:1;box-shadow:0 9px 20px rgba(181,31,37,.25);z-index:1;transition:transform .22s ease}
.indicador-card-link:hover .indicador-card-simbolo{transform:scale(1.06)}
.indicador-card-titulo{display:flex;align-items:center;justify-content:center;text-align:center;min-height:54px;margin:16px 0 0;font-size:1rem;font-weight:850;line-height:1.25;z-index:1}
@media(max-width:900px){.portal-brand{gap:12px;flex-wrap:wrap}.portal-title{font-size:23px;white-space:normal}.portal-logo-claro{width:68px;height:68px}.portal-logo-logistica{width:118px;height:56px}.portal-sep{height:42px}.indicador-card-link{min-height:158px;padding:21px 14px 18px}.indicador-card-titulo{font-size:.92rem}.indicador-card-simbolo{width:54px;height:54px;border-radius:17px;font-size:26px}}
@media(max-width:580px){.portal-head{padding:14px 12px}.portal-brand{display:grid;grid-template-columns:auto 1px minmax(160px,1fr) 1px auto;gap:8px}.portal-title{font-size:17px}.portal-logo-claro{width:52px;height:52px}.portal-logo-logistica{width:88px;height:44px}.portal-sep{height:34px}.portal-subtitle{font-size:11px}}




/* CONTROLES OPERACIONAIS - MODELO UNICO E LIMPO */
.block-container{padding-top:3.5rem!important}
.portal-head{min-height:126px!important;padding:20px 26px 16px!important;margin:-.25rem -1.25rem 18px!important;border-radius:0 0 24px 24px!important}
.portal-brand{min-height:88px!important;padding-top:0!important}

/* Linha abaixo da barra: areas na esquerda, usuario e sair na direita */
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor){display:grid!important;grid-template-columns:minmax(480px,600px) 1fr 190px!important;align-items:start!important;column-gap:18px!important;width:100%!important;margin:0 0 18px!important}
.menu-model-anchor,.exit-model-anchor{display:none!important}
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor)>div:nth-child(1){width:100%!important;min-width:0!important;max-width:none!important}
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor)>div:nth-child(2){width:100%!important;min-width:0!important}
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor)>div:nth-child(3){width:190px!important;min-width:190px!important;max-width:190px!important;justify-self:end!important}

/* Armazenagem, Triagem e Reversa exatamente como o segmented control do login */
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"]{width:100%!important;display:flex!important;justify-content:flex-start!important}
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"]>div{display:grid!important;grid-template-columns:repeat(3,minmax(0,1fr))!important;gap:0!important;width:100%!important;min-height:40px!important;background:#fff!important;border:1px solid #c8c8c8!important;border-radius:10px!important;overflow:hidden!important;box-shadow:none!important}
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] button,
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] label{width:100%!important;min-width:0!important;max-width:none!important;height:40px!important;min-height:40px!important;margin:0!important;padding:0 18px!important;border:0!important;border-right:1px solid #d0d0d0!important;border-radius:0!important;background:#fff!important;color:#111!important;box-shadow:none!important;transform:none!important;opacity:1!important}
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] button:last-child,
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] label:last-child{border-right:0!important}
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] button *,
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] label *{color:#111!important;font-size:1rem!important;font-weight:400!important;line-height:1!important;white-space:nowrap!important;overflow:visible!important;text-overflow:clip!important;opacity:1!important}
/* Selecionado igual ao Entrar: rosa claro, contorno e texto vermelho */
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] button[aria-checked="true"],
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] button[aria-pressed="true"],
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] button[data-state="on"],
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] label:has(input:checked){background:#fff1f1!important;color:#ff313b!important;box-shadow:inset 0 0 0 1px #ff313b!important;border-color:#ff313b!important}
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] button[aria-checked="true"] *,
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] button[aria-pressed="true"] *,
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] button[data-state="on"] *,
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] label:has(input:checked) *{color:#ff313b!important;font-weight:400!important;opacity:1!important}

/* Usuario e Sair. Sair no mesmo modelo, dimensoes e cores */
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) .header-user-name{text-align:center!important;color:#111!important;font-size:.95rem!important;font-weight:800!important;white-space:nowrap!important;padding:0 0 7px!important}
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) .st-key-btn_sair_header,
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) .st-key-btn_sair_header [data-testid="stButton"]{width:100%!important}
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) .st-key-btn_sair_header button{width:100%!important;height:40px!important;min-height:40px!important;margin:0!important;padding:0 18px!important;border:1px solid #c8c8c8!important;border-radius:10px!important;background:#fff!important;color:#111!important;box-shadow:none!important;transform:none!important}
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) .st-key-btn_sair_header button p,
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) .st-key-btn_sair_header button span{color:#111!important;font-size:1rem!important;font-weight:400!important;opacity:1!important}
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) .st-key-btn_sair_header button:hover{background:#fff1f1!important;border-color:#ff313b!important;box-shadow:inset 0 0 0 1px #ff313b!important}
div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) .st-key-btn_sair_header button:hover *{color:#ff313b!important}

/* Cabecalho compacto nas telas sem usuario */
body:not(:has(.menu-model-anchor)) .portal-head{min-height:126px!important;padding:20px 26px 16px!important}
body:not(:has(.menu-model-anchor)) .portal-brand{min-height:88px!important;padding-top:0!important}

@media(max-width:900px){
 .block-container{padding-top:3.2rem!important}
 div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor){grid-template-columns:1fr 190px!important;row-gap:10px!important}
 div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor)>div:nth-child(1){grid-column:1/2!important}
 div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor)>div:nth-child(2){display:none!important}
 div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor)>div:nth-child(3){grid-column:2/3!important}
}
@media(max-width:580px){
 .block-container{padding-top:3rem!important}
 .portal-head{min-height:158px!important;padding:18px 10px 16px!important}
 div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor){grid-template-columns:1fr!important}
 div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor)>div:nth-child(1){grid-column:1!important;width:100%!important}
 div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor)>div:nth-child(3){grid-column:1!important;width:160px!important;min-width:160px!important;justify-self:end!important}
 div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] button,
 div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] label{padding:0 5px!important;height:40px!important;min-height:40px!important}
 div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] button *,
 div[data-testid="stHorizontalBlock"]:has(.menu-model-anchor) [data-testid="stSegmentedControl"] label *{font-size:.78rem!important}
}

/* ASSETS EXTERNOS E RESPONSIVIDADE ESTAVEL */
.access-menu-anchor{display:none!important}
html,body,.stApp,[data-testid="stAppViewContainer"]{max-width:100%!important;overflow-x:hidden!important}
@media(max-width:900px){
 .block-container{padding-top:2.4rem!important;padding-left:.9rem!important;padding-right:.9rem!important}
 .portal-head,body:not(:has(.menu-model-anchor)) .portal-head{min-height:0!important;height:auto!important;width:100%!important;max-width:100%!important;margin:0 0 14px!important;padding:14px 16px!important;box-sizing:border-box!important;overflow:hidden!important}
 .portal-brand,body:not(:has(.menu-model-anchor)) .portal-brand{display:grid!important;grid-template-columns:70px 1px minmax(0,1fr) 1px 110px!important;align-items:center!important;justify-items:center!important;gap:10px!important;width:100%!important;max-width:100%!important;min-height:72px!important;height:auto!important;padding:0!important;margin:0!important;box-sizing:border-box!important;flex-wrap:nowrap!important}
 .portal-logo-claro{grid-column:1!important;width:66px!important;height:66px!important;max-width:66px!important;object-fit:contain!important}
 .portal-brand>.portal-sep:nth-of-type(1){grid-column:2!important;height:40px!important}
 .portal-title{grid-column:3!important;min-width:0!important;width:100%!important;max-width:100%!important;margin:0!important;padding:0 4px!important;text-align:center!important;font-size:clamp(1.12rem,3.2vw,1.55rem)!important;line-height:1.12!important;white-space:normal!important;word-break:normal!important;overflow-wrap:normal!important;writing-mode:horizontal-tb!important}
 .portal-brand>.portal-sep:nth-of-type(2){grid-column:4!important;height:40px!important}
 .portal-logo-logistica{grid-column:5!important;width:104px!important;height:54px!important;max-width:104px!important;object-fit:contain!important}
}
@media(max-width:580px){
 .block-container{padding-top:2.15rem!important;padding-left:.5rem!important;padding-right:.5rem!important}
 .portal-head,body:not(:has(.menu-model-anchor)) .portal-head{min-height:0!important;height:auto!important;margin:0 0 12px!important;padding:10px 8px!important;border-radius:0 0 16px 16px!important}
 .portal-brand,body:not(:has(.menu-model-anchor)) .portal-brand{display:grid!important;grid-template-columns:50px 1px minmax(0,1fr) 1px 72px!important;align-items:center!important;justify-items:center!important;gap:5px!important;min-height:58px!important;height:auto!important;width:100%!important;max-width:100%!important;padding:0!important;margin:0!important}
 .portal-logo-claro{grid-column:1!important;width:48px!important;height:48px!important;max-width:48px!important}
 .portal-brand>.portal-sep:nth-of-type(1){grid-column:2!important;width:1px!important;height:30px!important}
 .portal-title{grid-column:3!important;width:100%!important;min-width:0!important;max-width:100%!important;padding:0 2px!important;font-size:clamp(.72rem,3.55vw,.94rem)!important;line-height:1.1!important;white-space:normal!important;word-break:normal!important;overflow-wrap:normal!important;writing-mode:horizontal-tb!important;text-align:center!important}
 .portal-brand>.portal-sep:nth-of-type(2){grid-column:4!important;width:1px!important;height:30px!important}
 .portal-logo-logistica{grid-column:5!important;width:69px!important;height:40px!important;max-width:69px!important}
 .st-key-modo [data-testid="stSegmentedControl"]>div,.st-key-modo [data-testid="stSegmentedControl"]>div>div{display:grid!important;grid-template-columns:repeat(2,minmax(0,1fr))!important;grid-auto-flow:row!important;gap:6px!important;width:100%!important;max-width:100%!important;border:0!important;background:transparent!important;overflow:visible!important}
 .st-key-modo [data-testid="stSegmentedControl"] button,.st-key-modo [data-testid="stSegmentedControl"] label{width:100%!important;min-width:0!important;max-width:none!important;min-height:44px!important;height:auto!important;margin:0!important;padding:7px 5px!important;border:1px solid #c8c8c8!important;border-radius:10px!important;background:#fff!important}
 .st-key-modo [data-testid="stSegmentedControl"] button *,.st-key-modo [data-testid="stSegmentedControl"] label *{font-size:.78rem!important;line-height:1.12!important;white-space:normal!important;text-align:center!important}
}
</style>
<div class="portal-head portal-head-base">
  <div class="portal-brand">
    <img class="portal-logo-claro" src="logo_claro.png" alt="Claro">
    <span class="portal-sep"></span>
    <div class="portal-title">Central de Inteligência Operacional</div>
    <span class="portal-sep"></span>
    <img class="portal-logo-logistica" src="logo_logistica.png" alt="Logística">
  </div>
</div>
""", unsafe_allow_html=True)

if "modo" not in st.session_state:
    st.session_state.modo = "Entrar"
if "usuario_logado" not in st.session_state:
    st.session_state.usuario_logado = None
if "admin_logado" not in st.session_state:
    st.session_state.admin_logado = None

if st.session_state.admin_logado:
    with conectar() as con:
        pendentes = [dict(x) for x in con.execute("SELECT * FROM solicitacoes WHERE status='PENDENTE' ORDER BY criado_em").fetchall()]
    usuarios_admin = listar_usuarios()

    st.title("Administração")
    st.markdown(f'<div class="pendencia">{len(pendentes)} solicitação(ões) pendente(s)</div>', unsafe_allow_html=True)
    st.caption("Datas e horários apresentados no fuso de Brasília (America/Sao_Paulo).")
    st.write("")

    aba_pendencias, aba_usuarios, aba_importacao = st.tabs([
        f"Solicitações pendentes ({len(pendentes)})",
        f"Usuários cadastrados ({len(usuarios_admin)})",
        "Importar usuários",
    ])

    with aba_pendencias:
        if not pendentes:
            st.info("Não existem solicitações pendentes.")
        for item in pendentes:
            titulo = "Novo cadastro" if item["tipo"] == "CADASTRO" else "Redefinição de senha"
            with st.container(border=True):
                st.subheader(f"{titulo} | {item['usuario']}")
                st.write(f"**Nome:** {item['nome_completo'] or 'Cadastro legado'}")
                st.write(f"**E-mail:** {item['email']}")
                st.write(f"**Solicitado em:** {data_hora_brasilia(item['criado_em'])} (Brasília)")
                obs = st.text_input("Observação", key=f"obs_{item['id']}")
                c1, c2 = st.columns(2)
                if c1.button("Aprovar", key=f"aprovar_{item['id']}", type="primary", use_container_width=True):
                    decidir_solicitacao(item["id"], True, st.session_state.admin_logado, obs)
                    st.success(f"Solicitação aprovada em {data_hora_brasilia(agora_iso())} (Brasília).")
                    st.rerun()
                if c2.button("Rejeitar", key=f"rejeitar_{item['id']}", use_container_width=True):
                    decidir_solicitacao(item["id"], False, st.session_state.admin_logado, obs)
                    st.warning(f"Solicitação rejeitada em {data_hora_brasilia(agora_iso())} (Brasília).")
                    st.rerun()

    with aba_usuarios:
        if not usuarios_admin:
            st.info("Nenhum usuário aprovado foi cadastrado ainda.")
        for usuario_item in usuarios_admin:
            nome_exibicao = usuario_item["nome_completo"] or usuario_item["usuario"] or "Cadastro legado"
            status_atual = "Ativo" if usuario_item["ativo"] else "Desativado"
            with st.expander(f"{nome_exibicao} | {status_atual}", expanded=False):
                c1, c2 = st.columns(2)
                with c1:
                    st.write(f"**Nome:** {usuario_item['nome_completo'] or 'Não disponível no cadastro legado'}")
                    st.write(f"**Usuário:** {usuario_item['usuario'] or 'Não disponível no cadastro legado'}")
                    st.write(f"**E-mail:** {usuario_item['email'] or 'Não disponível no cadastro legado'}")
                    st.write(f"**Perfil:** {usuario_item['perfil']}")
                with c2:
                    st.write(f"**Status atual:** {status_atual}")
                    st.write(f"**Criado em:** {data_hora_brasilia(usuario_item['criado_em'])}")
                    st.write(f"**Aprovado em:** {data_hora_brasilia(usuario_item['aprovado_em'])}")
                    st.write(f"**Atualizado em:** {data_hora_brasilia(usuario_item['atualizado_em'])}")

                with st.expander("Editar cadastro", expanded=False):
                    st.caption(
                        "A alteração do usuário ou do e-mail atualizará os hashes utilizados no próximo login. "
                        "A senha atual será preservada."
                    )
                    with st.form(f"editar_usuario_{usuario_item['id']}"):
                        nome_editado = st.text_input(
                            "Nome completo",
                            value=usuario_item["nome_completo"] or "",
                            key=f"editar_nome_{usuario_item['id']}",
                        )
                        usuario_editado = st.text_input(
                            "Nome de usuário",
                            value=usuario_item["usuario"] or "",
                            key=f"editar_login_{usuario_item['id']}",
                        )
                        email_editado = st.text_input(
                            "E-mail corporativo",
                            value=usuario_item["email"] or "",
                            key=f"editar_email_{usuario_item['id']}",
                        )
                        confirmar_edicao = st.checkbox(
                            "Confirmo a atualização dos dados de acesso.",
                            key=f"confirmar_edicao_{usuario_item['id']}",
                        )
                        salvar_edicao = st.form_submit_button(
                            "Salvar alterações",
                            type="primary",
                            use_container_width=True,
                        )
                    if salvar_edicao:
                        if not confirmar_edicao:
                            st.warning("Marque a confirmação antes de salvar as alterações.")
                        else:
                            try:
                                editar_cadastro_usuario(
                                    usuario_item["id"],
                                    nome_editado,
                                    usuario_editado,
                                    email_editado,
                                    st.session_state.admin_logado,
                                )
                                st.success("Cadastro atualizado. O próximo login deverá usar os dados informados.")
                                st.rerun()
                            except ValueError as exc:
                                st.error(str(exc))

                acao1, acao2 = st.columns(2)
                if usuario_item["ativo"]:
                    if acao1.button("Desativar usuário", key=f"desativar_{usuario_item['id']}", use_container_width=True):
                        alterar_status_usuario(usuario_item["id"], False, st.session_state.admin_logado)
                        st.rerun()
                else:
                    if acao1.button("Reativar usuário", key=f"reativar_{usuario_item['id']}", type="primary", use_container_width=True):
                        alterar_status_usuario(usuario_item["id"], True, st.session_state.admin_logado)
                        st.rerun()

                confirmar_exclusao = acao2.checkbox("Confirmar exclusão", key=f"confirma_exclusao_{usuario_item['id']}")
                if acao2.button("Excluir definitivamente", key=f"excluir_{usuario_item['id']}", disabled=not confirmar_exclusao, use_container_width=True):
                    excluir_usuario(usuario_item["id"], st.session_state.admin_logado)
                    st.rerun()

    with aba_importacao:
        st.subheader("Importar usuários do JSON legado")
        st.info(
            "A importação inclui usuários novos e atualiza cadastros já existentes pelo hash do usuário ou do e-mail. "
            "Os demais usuários do banco serão preservados."
        )
        arquivo_json = st.file_uploader(
            "Selecione o usuarios.json antigo",
            type=["json"],
            accept_multiple_files=False,
            key="importar_usuarios_json",
        )
        confirmar_importacao = st.checkbox(
            "Confirmo que este arquivo contém usuários autorizados para este portal.",
            key="confirmar_importacao_json",
        )
        if st.button(
            "Importar usuários",
            type="primary",
            use_container_width=True,
            disabled=arquivo_json is None or not confirmar_importacao,
            key="btn_importar_usuarios_json",
        ):
            try:
                resultado = importar_usuarios_json_manual(arquivo_json, st.session_state.admin_logado)
                st.success(
                    f"Importação concluída. Analisados: {resultado['analisados']} | "
                    f"Incluídos: {resultado['incluidos']} | Atualizados: {resultado['atualizados']} | "
                    f"Ignorados: {resultado['ignorados']}."
                )
                if resultado["erros"]:
                    with st.expander("Registros ignorados"):
                        for erro in resultado["erros"]:
                            st.write(f"- {erro}")
                st.session_state["resultado_ultima_importacao"] = resultado
            except ValueError as exc:
                st.error(str(exc))

        st.divider()
        st.subheader("Completar nomes dos cadastros legados em lote")
        st.info(
            "Envie uma planilha XLSX com a guia 'Usuarios' e as colunas nome_completo, usuario e email. "
            "A atualização ocorre quando usuário e e-mail correspondem ao cadastro legado. "
            "Se nome, usuário ou e-mail já existirem, os valores serão substituídos pelos dados atuais da planilha."
        )
        arquivo_xlsx = st.file_uploader(
            "Selecione a Relação Usuários da Central.xlsx",
            type=["xlsx"], accept_multiple_files=False, key="complementar_usuarios_xlsx",
        )
        confirmar_xlsx = st.checkbox(
            "Confirmo que os dados correspondem aos usuários já importados.",
            key="confirmar_complementacao_xlsx",
        )
        if st.button(
            "Completar cadastros legados", type="primary", use_container_width=True,
            disabled=arquivo_xlsx is None or not confirmar_xlsx,
            key="btn_complementar_usuarios_xlsx",
        ):
            try:
                resultado_xlsx = complementar_cadastros_xlsx(arquivo_xlsx, st.session_state.admin_logado)
                st.success(
                    f"Complementação concluída. Analisados: {resultado_xlsx['analisados']} | "
                    f"Complementados: {resultado_xlsx['complementados']} | "
                    f"Já completos: {resultado_xlsx['ja_completos']} | "
                    f"Não localizados: {resultado_xlsx['nao_localizados']} | "
                    f"Ignorados: {resultado_xlsx['ignorados']}."
                )
                if resultado_xlsx["erros"]:
                    with st.expander("Linhas não atualizadas"):
                        for erro in resultado_xlsx["erros"]:
                            st.write(f"- {erro}")
            except ValueError as exc:
                st.error(str(exc))

    st.download_button("Exportar usuarios.json compatível", exportar_json_compatibilidade(), "usuarios.json", "application/json")
    if st.button("Sair da administração"):
        st.session_state.admin_logado = None
        st.rerun()
    st.stop()

if st.session_state.usuario_logado:
    exibir_central_indicadores()
    st.stop()

opcoes = ["Entrar", "Solicitar cadastro", "Esqueci minha senha", "Administrador"]
st.markdown('<span class="access-menu-anchor"></span>', unsafe_allow_html=True)
st.segmented_control(
    "Acesso",
    opcoes,
    key="modo",
    label_visibility="collapsed",
)

if st.session_state.modo == "Entrar":
    st.subheader("Acesso à Central")
    with st.form("login"):
        usuario = st.text_input("Nome de usuário")
        email = st.text_input("E-mail corporativo")
        senha = st.text_input("Senha", type="password")
        enviar = st.form_submit_button("Entrar", type="primary", use_container_width=True)
    if enviar:
        u = autenticar_usuario(usuario, email, senha)
        if u:
            st.session_state.usuario_logado = u
            st.rerun()
        st.error("Dados inválidos ou acesso ainda não aprovado.")

elif st.session_state.modo == "Solicitar cadastro":
    st.subheader("Solicitação de cadastro")
    with st.form("cadastro"):
        nome = st.text_input("Nome completo")
        usuario = st.text_input("Nome de usuário")
        email = st.text_input("E-mail corporativo")
        senha = st.text_input("Senha", type="password")
        confirmar = st.text_input("Confirmar senha", type="password")
        enviar = st.form_submit_button("Enviar para aprovação", type="primary", use_container_width=True)
    if enviar:
        try:
            if senha != confirmar:
                raise ValueError("A confirmação da senha não corresponde.")
            token = criar_solicitacao_cadastro(nome, usuario, email, senha)
            st.success("Solicitação registrada. O acesso será liberado somente após aprovação do administrador.")
            st.code(token, language=None)
            st.caption("Guarde este código para ativar um novo navegador.")
        except ValueError as e:
            st.error(str(e))

elif st.session_state.modo == "Esqueci minha senha":
    st.subheader("Solicitar redefinição de senha")
    with st.form("redefinir"):
        usuario = st.text_input("Nome de usuário")
        email = st.text_input("E-mail corporativo")
        senha = st.text_input("Nova senha", type="password")
        confirmar = st.text_input("Confirmar nova senha", type="password")
        enviar = st.form_submit_button("Enviar para aprovação", type="primary", use_container_width=True)
    if enviar:
        if senha != confirmar:
            st.error("A confirmação da nova senha não corresponde.")
        else:
            try:
                criar_solicitacao_senha(usuario, email, senha)
                st.success("Se os dados corresponderem a um cadastro ativo, a solicitação será encaminhada para aprovação.")
            except ValueError as e:
                st.error(str(e))

else:
    st.subheader("Acesso administrativo")
    with st.form("admin"):
        usuario = st.text_input("Usuário administrador")
        senha = st.text_input("Senha administrativa", type="password")
        enviar = st.form_submit_button("Acessar administração", type="primary", use_container_width=True)
    if enviar:
        if credenciais_admin_validas(usuario, senha):
            st.session_state.admin_logado = usuario
            st.rerun()
        st.error("Credenciais administrativas inválidas ou não configuradas.")
