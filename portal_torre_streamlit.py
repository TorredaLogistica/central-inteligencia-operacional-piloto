import hashlib
import hmac
import json
import os
import secrets
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

import streamlit as st

DB_PATH = Path(os.getenv("TORRE_DB_PATH", "torre_usuarios.db"))
USUARIOS_JSON = Path(os.getenv("TORRE_USUARIOS_JSON", "usuarios.json"))
ITERACOES = 600_000

st.set_page_config(page_title="Claro | Central de Inteligência Operacional", page_icon="🔴", layout="wide")


def agora_iso():
    return datetime.now(timezone.utc).isoformat()


def normalizar_usuario(valor: str) -> str:
    import unicodedata
    valor = unicodedata.normalize("NFD", valor or "")
    valor = "".join(c for c in valor if unicodedata.category(c) != "Mn")
    return "".join(c for c in valor.upper() if "A" <= c <= "Z")


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


def exportar_json_compatibilidade():
    with conectar() as con:
        linhas = con.execute("SELECT nome_hash,email_hash,senha_hash,salt,iteracoes,codigo_dispositivo_hash,ativo,aprovado_em FROM usuarios").fetchall()
    return json.dumps({"versao": 28, "usuarios": [dict(x) for x in linhas], "atualizado_em": agora_iso()}, ensure_ascii=False, indent=2)


iniciar_banco()
importar_json_legado()

st.markdown("""
<style>
.stApp{background:#eef0f3}.claro-head{padding:18px 24px;border-radius:0 0 18px 18px;background:linear-gradient(180deg,#b51f25,#f47b45);color:white;margin:-1rem -1rem 1.5rem}.claro-head h1{margin:0;font-size:1.65rem}.claro-head p{margin:.35rem 0 0;opacity:.9}.pendencia{padding:.8rem 1rem;border-radius:12px;background:#fff3cd;border:1px solid #ffec99;color:#7a5200;font-weight:700}
</style><div class="claro-head"><h1>Claro | Central de Inteligência Operacional</h1><p>Cadastro, aprovação e recuperação de acesso</p></div>
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
    st.title("Administração")
    st.markdown(f'<div class="pendencia">{len(pendentes)} solicitação(ões) pendente(s)</div>', unsafe_allow_html=True)
    st.write("")
    for item in pendentes:
        titulo = "Novo cadastro" if item["tipo"] == "CADASTRO" else "Redefinição de senha"
        with st.container(border=True):
            st.subheader(f"{titulo} | {item['usuario']}")
            st.write(f"**Nome:** {item['nome_completo'] or 'Cadastro legado'}")
            st.write(f"**E-mail:** {item['email']}")
            st.write(f"**Solicitado em:** {item['criado_em']}")
            obs = st.text_input("Observação", key=f"obs_{item['id']}")
            c1, c2 = st.columns(2)
            if c1.button("Aprovar", key=f"aprovar_{item['id']}", type="primary", use_container_width=True):
                decidir_solicitacao(item["id"], True, st.session_state.admin_logado, obs)
                st.rerun()
            if c2.button("Rejeitar", key=f"rejeitar_{item['id']}", use_container_width=True):
                decidir_solicitacao(item["id"], False, st.session_state.admin_logado, obs)
                st.rerun()
    st.download_button("Exportar usuarios.json compatível", exportar_json_compatibilidade(), "usuarios.json", "application/json")
    if st.button("Sair da administração"):
        st.session_state.admin_logado = None
        st.rerun()
    st.stop()

if st.session_state.usuario_logado:
    st.success("Acesso aprovado e autenticação concluída.")
    st.write("A área dos indicadores pode ser incorporada neste ponto ou permanecer no portal HTML durante a migração.")
    if st.button("Sair"):
        st.session_state.usuario_logado = None
        st.rerun()
    st.stop()

opcoes = ["Entrar", "Solicitar cadastro", "Esqueci minha senha", "Administrador"]
st.session_state.modo = st.segmented_control("Acesso", opcoes, default=st.session_state.modo) or "Entrar"

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
