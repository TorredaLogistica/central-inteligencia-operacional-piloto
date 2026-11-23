import hashlib
import hmac
import json
import io
import requests
import html
import os
import secrets
import sqlite3
from urllib.parse import quote, urlparse
from datetime import datetime, timezone, timedelta
from zoneinfo import ZoneInfo
from pathlib import Path
from io import BytesIO

import streamlit as st
import pandas as pd
import altair as alt
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from streamlit.components.v1 import html as components_html

DB_PATH = Path(os.getenv("TORRE_DB_PATH", "torre_usuarios.db"))
USUARIOS_JSON = Path(os.getenv("TORRE_USUARIOS_JSON", "usuarios.json"))
ITERACOES = 600_000
FUSO_BRASILIA = ZoneInfo("America/Sao_Paulo")

# Catálogo dos indicadores. Para disponibilizar um indicador, preencha a URL.
# Os títulos foram mantidos alinhados à Central HTML atual.
INDICADORES = {
    "Armazenagem": [
        {"titulo": "Chatbot Torre Logística", "url": "https://chatbot-logistica-e4cpfbcta3qyqchopsdjeg.streamlit.app/", "icone": "🤖", "periodicidade": "Em Construção", "ultima_atualizacao": "-"},
        {"titulo": "Separação e Faturamento", "url": "https://dashboard-slaseparacaofaturamento-mdnfzinkaebwzysne83ewp.streamlit.app/", "icone": "📊", "periodicidade": "Esporádica", "ultima_atualizacao": "15/11/2026"},
        {"titulo": "Pedidos para LPs", "url": "https://pedidoslpsaas-y44bkbmcg4kon8fogbro34.streamlit.app/", "icone": "📦", "periodicidade": "Mensal", "ultima_atualizacao": "Outubro/2026"},
        {"titulo": "Resultado do DRE", "url": "https://resultadodre-lk6rh4ahefeuwfhwrg2ioc.streamlit.app/", "icone": "💰", "periodicidade": "Mensal", "ultima_atualizacao": "Outubro/2026"},
        {"titulo": "Valores dos EAs", "url": "https://valorestoques-eas-73bxfsks3rnoxjo44fuqm7.streamlit.app/", "icone": "📈", "periodicidade": "Mensal", "ultima_atualizacao": "Novembro/2026"},
        {"titulo": "Atendimento de OVs nos TLs", "url": "https://atendimento-de-ovs-nos-tls-in3rykeacnvjxedhb7r9zc.streamlit.app/", "icone": "📋", "periodicidade": "Esporádica", "ultima_atualizacao": "17/11/2026"},
        {"titulo": "Taxa de Ocupação dos CDs", "url": "https://taxadeocupacaodoscds-tfx8ftu78n46vhvn5dxc7k.streamlit.app/", "icone": "🏭", "periodicidade": "Esporádica", "ultima_atualizacao": "13/11/2026"},
        {"titulo": "Pedidos Canal Vermelho", "url": "https://cuencjwy3ahhnzymzsspuc.streamlit.app/", "icone": "⚡", "periodicidade": "Esporádica", "ultima_atualizacao": "15/11/2026"},
        {"titulo": "Pedidos LPs e AAs NFs não Agrupadas", "url": "https://nfs-nao-agrupadas-juqvjn8nhbzuhzzdfwknbl.streamlit.app/", "icone": "🧾", "periodicidade": "Esporádica", "ultima_atualizacao": "01/10 a 17/11/2026"},
        {"titulo": "Faturas dos OPLs", "url": "https://faturasdosopls-fdgzskwvbciwkgubekfjcz.streamlit.app/", "icone": "💵", "periodicidade": "Mensal", "ultima_atualizacao": "Novembro/2026"},
        {"titulo": "Recebimento de Usados", "url": "https://controlederecebimentodeusados-ucfnrvqwuceiet5q7tt4wn.streamlit.app/", "icone": "♻️", "periodicidade": "Mensal", "ultima_atualizacao": "Outubro/2026"},
        {"titulo": "Nível de Serviços dos OPLs", "url": "https://niveldeservicoopls-sgrryyugyheukmxp2xtzp8.streamlit.app/", "icone": "🚚", "periodicidade": "Esporádica", "ultima_atualizacao": "15/11/2026"},
        {"titulo": "Forecast e Realizado", "url": "https://forecasterealizado-kebvdtavq5yc8s9kwrfqiu.streamlit.app/", "icone": "🔢", "periodicidade": "Mensal", "ultima_atualizacao": "Novembro/2026"},
        {"titulo": "Simulação de Pedidos", "url": "https://simulacaopedidos-myjtjrm3nklxzprutjpbd5.streamlit.app/", "icone": "🧮", "periodicidade": "Mensal", "ultima_atualizacao": "Outubro/2026"},
    ],
    "Triagem": [],
    "Reversa": [],
}

FONTES_REMOTAS = {
    "Separação e Faturamento": {"repositorio": "TorredaLogistica/dashboard-sla_separacao_faturamento", "branch": "main", "arquivo": "Faturamento SLA 2026.xlsb", "formato": "xlsb"},
    "Pedidos para LPs": {"repositorio": "TorredaLogistica/Pedidos_LPs_AAs", "branch": "main", "arquivo": "Base OTIF SOE AAs e LPs.xlsx", "formato": "xlsx"},
    "Resultado do DRE": {"repositorio": "TorredaLogistica/Resultado_DRE", "branch": "main", "arquivo": "Resultado DRE.xlsx", "formato": "xlsx"},
    "Valores dos EAs": {"repositorio": "TorredaLogistica/Valor_Estoques_-EAs", "branch": "main", "arquivo": "Valor Estoque EAs_Consolidado.parquet", "formato": "parquet"},
    "Atendimento de OVs nos TLs": {"repositorio": "TorredaLogistica/Atendimento-de-OVs-nos-TLs", "branch": "main", "arquivo": "Base OVs TLs.parquet", "formato": "parquet"},
    "Taxa de Ocupação dos CDs": {"repositorio": "TorredaLogistica/Taxa_de_Ocupacao_dos_CDs", "branch": "main", "arquivo": "Ocupação dos CDs.xlsx", "formato": "xlsx"},
    "Pedidos Canal Vermelho": {"repositorio": "TorredaLogistica/Canal_Vermelho", "branch": "main", "arquivo": "Base OTIF.parquet", "formato": "parquet"},
    "Pedidos LPs e AAs NFs não Agrupadas": {"repositorio": "TorredaLogistica/NFs-Nao-Agrupadas", "branch": "main", "arquivo": "Base OTIF SOE AAs e LPs.xlsx", "formato": "xlsx"},
    "Faturas dos OPLs": {"repositorio": "TorredaLogistica/Faturas_dos_OPLs", "branch": "main", "arquivo": "Base Fatura.xlsx", "formato": "xlsx"},
    "Recebimento de Usados": {"repositorio": "TorredaLogistica/Controle_de_Recebimento_de_Usados", "branch": "main", "arquivo": "Controle de Recebimento de Materiais.xlsx", "formato": "xlsx"},
    "Nível de Serviços dos OPLs": {"repositorio": "TorredaLogistica/Nivel_De_Servico_OPLs", "branch": "main", "arquivo": "Base Otif 2026.parquet", "formato": "parquet"},
    "Forecast e Realizado": {"repositorio": "TorredaLogistica/Forecast_e_Realizado", "branch": "main", "arquivo": "Farol.xlsx", "formato": "xlsx"},
    "Simulação de Pedidos": {"repositorio": "TorredaLogistica/Simulacao_Pedidos", "branch": "main", "arquivo": "Base_Pedidos_AAs_PLs_2026.parquet", "formato": "parquet"},
}

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


def registrar_acesso_indicador(usuario_atual, area, indicador, url):
    if "sessao_acesso_id" not in st.session_state:
        st.session_state.sessao_acesso_id = secrets.token_hex(16)
    with conectar() as con:
        con.execute(
            """INSERT INTO acessos_indicadores
               (usuario_id,nome_completo,usuario,area,indicador,url,acessado_em,sessao_id)
               VALUES(?,?,?,?,?,?,?,?)""",
            (
                usuario_atual.get("id"),
                usuario_atual.get("nome_completo"),
                usuario_atual.get("usuario"),
                area,
                indicador,
                url,
                agora_iso(),
                st.session_state.sessao_acesso_id,
            ),
        )


def excluir_acessos_indicadores(ids_acessos, administrador, motivo="EXCLUSAO_SELECIONADA"):
    ids_validos = sorted({int(acesso_id) for acesso_id in ids_acessos if acesso_id is not None})
    if not ids_validos:
        return 0
    placeholders = ",".join("?" for _ in ids_validos)
    with conectar() as con:
        registros = con.execute(
            f"SELECT * FROM acessos_indicadores WHERE id IN ({placeholders})",
            ids_validos,
        ).fetchall()
        if not registros:
            return 0
        referencia = json.dumps(
            {
                "motivo": motivo,
                "quantidade": len(registros),
                "ids": [registro["id"] for registro in registros],
            },
            ensure_ascii=False,
        )
        con.execute(f"DELETE FROM acessos_indicadores WHERE id IN ({placeholders})", ids_validos)
        con.execute(
            "INSERT INTO auditoria(acao,referencia,administrador,criado_em) VALUES(?,?,?,?)",
            ("EXCLUIR_ACESSOS_INDICADORES", referencia, administrador, agora_iso()),
        )
    return len(registros)


def consultar_acessos_indicadores(data_inicial, data_final, area=None, indicador=None, usuario_id=None):
    inicio_local = datetime.combine(data_inicial, datetime.min.time(), tzinfo=FUSO_BRASILIA)
    fim_local = datetime.combine(data_final, datetime.max.time(), tzinfo=FUSO_BRASILIA)
    inicio_utc = inicio_local.astimezone(timezone.utc).isoformat()
    fim_utc = fim_local.astimezone(timezone.utc).isoformat()
    filtros = ["acessado_em BETWEEN ? AND ?"]
    parametros = [inicio_utc, fim_utc]
    if area and area != "Todas":
        filtros.append("area=?")
        parametros.append(area)
    if indicador and indicador != "Todos":
        filtros.append("indicador=?")
        parametros.append(indicador)
    if usuario_id:
        filtros.append("usuario_id=?")
        parametros.append(usuario_id)
    with conectar() as con:
        linhas = con.execute(
            f"""SELECT id,usuario_id,nome_completo,usuario,area,indicador,url,acessado_em,sessao_id
                FROM acessos_indicadores
                WHERE {' AND '.join(filtros)}
                ORDER BY acessado_em DESC""",
            parametros,
        ).fetchall()
    return [dict(linha) for linha in linhas]


def exibir_relatorio_acessos(usuarios_admin):
    st.subheader("Acessos aos indicadores")
    st.caption(
        "O relatório registra somente a abertura do indicador pela Central. Não registra filtros, conteúdo consultado "
        "ou ações realizadas dentro do indicador. Utilize os dados para análise de adoção e suporte, não para avaliação individual de desempenho."
    )
    hoje = datetime.now(FUSO_BRASILIA).date()
    primeiro_dia = hoje.replace(day=1)
    f1, f2, f3 = st.columns(3)
    data_inicial = f1.date_input("Data inicial", value=primeiro_dia, key="acesso_data_inicial")
    data_final = f2.date_input("Data final", value=hoje, key="acesso_data_final")
    area = f3.selectbox("Área", ["Todas"] + list(INDICADORES.keys()), key="acesso_area")

    nomes_indicadores = sorted({item["titulo"] for itens in INDICADORES.values() for item in itens})
    usuarios_opcoes = {"Todos": None}
    for item in usuarios_admin:
        rotulo = item.get("nome_completo") or item.get("usuario") or f"Cadastro {item['id']}"
        usuarios_opcoes[f"{rotulo} | {item.get('usuario') or 'legado'}"] = item["id"]
    f4, f5 = st.columns(2)
    indicador = f4.selectbox("Indicador", ["Todos"] + nomes_indicadores, key="acesso_indicador")
    usuario_rotulo = f5.selectbox("Usuário", list(usuarios_opcoes.keys()), key="acesso_usuario")

    if data_inicial > data_final:
        st.error("A data inicial não pode ser posterior à data final.")
        return

    acessos = consultar_acessos_indicadores(
        data_inicial, data_final, area, indicador, usuarios_opcoes[usuario_rotulo]
    )
    total = len(acessos)
    usuarios_distintos = len({x["usuario_id"] for x in acessos if x["usuario_id"] is not None})
    indicadores_distintos = len({x["indicador"] for x in acessos})
    ultimo = data_hora_brasilia(acessos[0]["acessado_em"]) if acessos else "-"
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Total de acessos", total)
    k2.metric("Usuários distintos", usuarios_distintos)
    k3.metric("Indicadores acessados", indicadores_distintos)
    k4.metric("Último acesso", ultimo)

    if not acessos:
        st.info("Nenhum acesso foi registrado para os filtros selecionados.")
    else:
        df = pd.DataFrame(acessos)
        df["Data e hora"] = df["acessado_em"].apply(data_hora_brasilia)
        df["Nome"] = df["nome_completo"].fillna("Cadastro legado")
        df["Usuário"] = df["usuario"].fillna("Não disponível")
        st.markdown("#### Detalhamento dos acessos")
        tabela_exclusao = df[["id", "Data e hora", "Nome", "Usuário", "area", "indicador"]].rename(
            columns={"area": "Área", "indicador": "Indicador"}
        )
        tabela_exclusao.insert(0, "Excluir", False)
        tabela_editada = st.data_editor(
            tabela_exclusao,
            use_container_width=True,
            hide_index=True,
            disabled=["id", "Data e hora", "Nome", "Usuário", "Área", "Indicador"],
            column_config={
                "Excluir": st.column_config.CheckboxColumn(
                    "Selecionar", help="Marque os registros que devem ser excluídos."
                ),
                "id": None,
            },
            key="grade_exclusao_acessos",
        )
        ids_selecionados = tabela_editada.loc[tabela_editada["Excluir"], "id"].astype(int).tolist()

        por_indicador = (
            df.groupby(["area", "indicador"], dropna=False)
              .agg(Acessos=("id", "count"), Usuarios_distintos=("usuario_id", "nunique"), Ultimo_acesso=("acessado_em", "max"))
              .reset_index()
        )
        por_indicador["Último acesso"] = por_indicador["Ultimo_acesso"].apply(data_hora_brasilia)
        st.markdown("#### Resumo por indicador")
        st.dataframe(
            por_indicador[["area", "indicador", "Acessos", "Usuarios_distintos", "Último acesso"]].rename(
                columns={"area": "Área", "indicador": "Indicador", "Usuarios_distintos": "Usuários distintos"}
            ),
            use_container_width=True,
            hide_index=True,
        )

    if acessos:
        st.markdown("#### Exclusão de registros")
        st.caption(
            "Marque os registros desejados na coluna Selecionar ou utilize a exclusão de todos os acessos "
            "retornados pelos filtros atuais. As exclusões ficam registradas na auditoria administrativa."
        )
        excluir_col1, excluir_col2 = st.columns(2)
        with excluir_col1:
            confirmar_selecionados = st.checkbox(
                f"Confirmo a exclusão dos registros selecionados ({len(ids_selecionados)}).",
                key="confirmar_exclusao_acessos_selecionados",
            )
            if st.button(
                "Excluir registros selecionados",
                use_container_width=True,
                disabled=not ids_selecionados or not confirmar_selecionados,
                key="excluir_acessos_selecionados",
            ):
                quantidade = excluir_acessos_indicadores(
                    ids_selecionados,
                    st.session_state.admin_logado,
                    motivo="EXCLUSAO_SELECIONADA",
                )
                st.success(f"{quantidade} registro(s) de acesso excluído(s).")
                st.rerun()
        with excluir_col2:
            ids_filtrados = [int(item["id"]) for item in acessos]
            confirmar_periodo = st.checkbox(
                f"Confirmo a exclusão de todos os acessos filtrados ({len(ids_filtrados)}).",
                key="confirmar_exclusao_todos_acessos_filtrados",
            )
            if st.button(
                "Excluir todos os acessos filtrados",
                use_container_width=True,
                disabled=not confirmar_periodo,
                key="excluir_todos_acessos_filtrados",
            ):
                quantidade = excluir_acessos_indicadores(
                    ids_filtrados,
                    st.session_state.admin_logado,
                    motivo="EXCLUSAO_TOTAL_FILTROS",
                )
                st.success(f"{quantidade} registro(s) de acesso excluído(s).")
                st.rerun()

    ids_com_acesso = {x["usuario_id"] for x in acessos if x["usuario_id"] is not None}
    sem_acesso = [
        {
            "Nome": u.get("nome_completo") or "Cadastro legado",
            "Usuário": u.get("usuario") or "Não disponível",
            "Status": "Ativo" if u.get("ativo") else "Desativado",
        }
        for u in usuarios_admin if u.get("ativo") and u["id"] not in ids_com_acesso
    ]
    with st.expander(f"Usuários ativos sem clique registrado no período ({len(sem_acesso)})"):
        st.caption("A ausência de registro significa apenas que não houve abertura pela Central no período selecionado.")
        if sem_acesso:
            st.dataframe(pd.DataFrame(sem_acesso), use_container_width=True, hide_index=True)
        else:
            st.success("Todos os usuários ativos possuem ao menos um clique registrado no período.")

    if acessos:
        csv = pd.DataFrame(acessos).to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Exportar acessos filtrados (CSV)",
            csv,
            f"acessos_indicadores_{data_inicial}_{data_final}.csv",
            "text/csv",
            use_container_width=True,
        )


def _github_token():
    try:
        return str(st.secrets.get("GITHUB_TOKEN", "")).strip()
    except Exception:
        return ""


def _github_headers():
    headers = {"Accept": "application/vnd.github+json", "X-GitHub-Api-Version": "2022-11-28"}
    token = _github_token()
    if token:
        headers["Authorization"] = f"Bearer {token}"
    return headers


@st.cache_data(ttl=1800, show_spinner=False)
def obter_metadados_fonte(nome, repositorio, branch, arquivo, token_marcador=""):
    caminho = quote(arquivo, safe="/")
    url = f"https://api.github.com/repos/{repositorio}/contents/{caminho}?ref={quote(branch)}"
    resposta = requests.get(url, headers=_github_headers(), timeout=25)
    if resposta.status_code == 404:
        raise FileNotFoundError(f"Arquivo não localizado em {repositorio}: {arquivo}")
    resposta.raise_for_status()
    dados = resposta.json()
    return {
        "Indicador": nome,
        "Repositório": repositorio,
        "Arquivo": arquivo,
        "Formato": Path(arquivo).suffix.lower().lstrip("."),
        "Tamanho (MB)": round(float(dados.get("size", 0)) / 1024 / 1024, 2),
        "Download": dados.get("download_url"),
        "SHA": dados.get("sha"),
        "Status": "Disponível",
    }


@st.cache_data(ttl=1800, show_spinner=False)
def carregar_base_github(nome, repositorio, branch, arquivo, formato, token_marcador=""):
    meta = obter_metadados_fonte(nome, repositorio, branch, arquivo, token_marcador)
    download_url = meta.get("Download")
    if not download_url:
        raise ValueError("O GitHub não retornou uma URL de download para a base.")
    resposta = requests.get(download_url, headers=_github_headers(), timeout=180)
    resposta.raise_for_status()
    conteudo = io.BytesIO(resposta.content)
    formato = str(formato).lower()
    if formato == "parquet":
        return pd.read_parquet(conteudo)
    if formato == "xlsx":
        return pd.read_excel(conteudo, engine="openpyxl")
    if formato == "xlsb":
        return pd.read_excel(conteudo, engine="calamine")
    raise ValueError(f"Formato remoto não suportado: {formato}")


def diagnosticar_base(df):
    datas_candidatas = []
    palavras_data = ("data", "dt ", "date", "emissão", "emissao", "criação", "criacao")
    for coluna in df.columns:
        if any(palavra in str(coluna).lower() for palavra in palavras_data):
            serie = pd.to_datetime(df[coluna], errors="coerce", dayfirst=True)
            if serie.notna().any(): datas_candidatas.append((str(coluna), serie.max()))
    coluna_data, data_maxima = (max(datas_candidatas, key=lambda x: x[1]) if datas_candidatas else (None, None))
    return {
        "Linhas": int(len(df)),
        "Colunas": int(len(df.columns)),
        "Coluna de data": coluna_data or "Não identificada",
        "Data máxima": data_maxima.strftime("%d/%m/%Y") if data_maxima is not None and not pd.isna(data_maxima) else "Não identificada",
    }


def exibir_fontes_dados():
    st.markdown('<div class="overview-title">Fontes de dados do GitHub</div>', unsafe_allow_html=True)
    st.caption("Monitoramento dos arquivos usados pelos indicadores. Os metadados são atualizados a cada 30 minutos.")
    c1, c2 = st.columns([1, 3])
    with c1:
        if st.button("Atualizar fontes agora", use_container_width=True, key="atualizar_fontes_remotas"):
            st.cache_data.clear(); st.rerun()
    with c2:
        st.caption("Para repositórios privados, configure GITHUB_TOKEN nos Secrets do Streamlit com permissão somente de leitura.")
    registros=[]
    token_marcador = "configurado" if _github_token() else "publico"
    for nome, fonte in FONTES_REMOTAS.items():
        try:
            registros.append(obter_metadados_fonte(nome, fonte["repositorio"], fonte["branch"], fonte["arquivo"], token_marcador))
        except Exception as exc:
            registros.append({"Indicador":nome,"Repositório":fonte["repositorio"],"Arquivo":fonte["arquivo"],"Formato":fonte["formato"],"Tamanho (MB)":None,"Download":None,"SHA":None,"Status":f"Indisponível: {str(exc)[:120]}"})
    df_status=pd.DataFrame(registros)
    disponiveis=int((df_status["Status"]=="Disponível").sum())
    k1,k2,k3=st.columns(3)
    k1.metric("Fontes configuradas",len(df_status)); k2.metric("Disponíveis",disponiveis); k3.metric("Com falha",len(df_status)-disponiveis)
    st.dataframe(df_status[["Indicador","Repositório","Arquivo","Formato","Tamanho (MB)","Status"]],use_container_width=True,hide_index=True)
    st.markdown("### Diagnóstico de uma base")
    selecionado=st.selectbox("Indicador",list(FONTES_REMOTAS.keys()),key="fonte_diagnostico")
    if st.button("Carregar e analisar base selecionada",type="primary",use_container_width=True,key="analisar_fonte_remota"):
        fonte=FONTES_REMOTAS[selecionado]
        try:
            with st.spinner(f"Baixando e analisando {selecionado}..."):
                df=carregar_base_github(selecionado,fonte["repositorio"],fonte["branch"],fonte["arquivo"],fonte["formato"],token_marcador)
            diagnostico=diagnosticar_base(df)
            cols=st.columns(4)
            for col,(rotulo,valor) in zip(cols,diagnostico.items()): col.metric(rotulo,valor)
            st.dataframe(df.head(50),use_container_width=True,hide_index=True)
        except Exception as exc:
            st.error(f"Não foi possível carregar a base: {exc}")


def catalogo_visao_geral():
    hoje = datetime.now(FUSO_BRASILIA).date()
    meses = {"janeiro":1,"fevereiro":2,"março":3,"abril":4,"maio":5,"junho":6,"julho":7,"agosto":8,"setembro":9,"outubro":10,"novembro":11,"dezembro":12}
    resultado=[]
    for area,itens in INDICADORES.items():
        for item in itens:
            periodicidade=item.get("periodicidade") or "A definir"
            ultima=item.get("ultima_atualizacao") or "-"
            data_ref=None
            try:
                if ultima.count("/")==2 and " a " not in ultima.lower(): data_ref=datetime.strptime(ultima,"%d/%m/%Y").date()
                else:
                    for nome,numero in meses.items():
                        if nome in ultima.lower():
                            ano=int(ultima.split("/")[-1]); data_ref=(datetime(ano+1,1,1)-timedelta(days=1)).date() if numero==12 else (datetime(ano,numero+1,1)-timedelta(days=1)).date(); break
            except (ValueError,TypeError): data_ref=None
            if "constru" in periodicidade.lower() or not item.get("url"): situacao="Em construção"
            elif data_ref and (hoje-data_ref).days>45: situacao="Revisar atualização"
            else: situacao="Disponível"
            resultado.append({"Área":area,"Indicador":item["titulo"],"Periodicidade":periodicidade,"Última atualização":ultima,"Situação":situacao,"Ícone":item.get("icone") or "📊"})
    return resultado


def card_executivo(titulo,valor,detalhe,classe=""):
    st.markdown(f'<div class="overview-kpi {classe}"><div class="overview-kpi-label">{html.escape(str(titulo))}</div><div class="overview-kpi-value">{html.escape(str(valor))}</div><div class="overview-kpi-detail">{html.escape(str(detalhe))}</div></div>',unsafe_allow_html=True)


def exibir_visao_geral():
    perfis_cd = {
        "Consolidado": {"volume":1280000,"expedicoes":1140000,"custo":12.47,"sla":96.8,"reversa":2.45,"fator":1.00},
        "Campinas": {"volume":238450,"expedicoes":219700,"custo":2.18,"sla":98.2,"reversa":1.88,"fator":0.19},
        "Rio de Janeiro": {"volume":205320,"expedicoes":186400,"custo":2.04,"sla":96.7,"reversa":2.38,"fator":0.16},
        "Brasília": {"volume":141880,"expedicoes":126300,"custo":1.42,"sla":96.1,"reversa":2.42,"fator":0.11},
        "Barueri": {"volume":186750,"expedicoes":169800,"custo":1.81,"sla":97.4,"reversa":2.05,"fator":0.15},
        "Contagem": {"volume":152460,"expedicoes":134900,"custo":1.55,"sla":95.6,"reversa":2.72,"fator":0.12},
        "Palhoça": {"volume":98740,"expedicoes":87900,"custo":1.02,"sla":95.2,"reversa":2.81,"fator":0.08},
        "Jaboatão": {"volume":88760,"expedicoes":77600,"custo":0.94,"sla":94.8,"reversa":2.96,"fator":0.07},
        "Manaus": {"volume":69340,"expedicoes":59400,"custo":0.82,"sla":93.9,"reversa":3.18,"fator":0.05},
        "Salvador": {"volume":98290,"expedicoes":87900,"custo":0.99,"sla":95.0,"reversa":2.89,"fator":0.08},
    }
    empresas = {
        "Consolidado":{"fator":1.00,"score":91.8}, "Claro Móvel":{"fator":0.48,"score":94.6},
        "Claro TV":{"fator":0.18,"score":90.1}, "Embratel":{"fator":0.14,"score":92.7},
        "NET":{"fator":0.12,"score":87.9}, "Claro Fixo":{"fator":0.08,"score":89.4},
    }
    t1,t2,t3,t4=st.columns([3.15,1.05,1.15,1.15])
    with t1:
        st.markdown('<div class="cio-title">Central de Inteligência Operacional</div><div class="cio-subtitle">Gerência de Armazenagem, Triagem e Reversa</div>',unsafe_allow_html=True)
    with t2:
        st.markdown('<div class="cio-update">📅 &nbsp; Última atualização:<br><strong>23/11/2026 12:30</strong></div>',unsafe_allow_html=True)
    with t3:
        cd_selecionado=st.selectbox("CD Origem",list(perfis_cd.keys()),key="cd_visao_geral")
    with t4:
        empresa_selecionada=st.selectbox("Empresa",list(empresas.keys()),key="empresa_visao_geral")

    fator_empresa=empresas[empresa_selecionada]["fator"]
    perfil=perfis_cd[cd_selecionado].copy()
    perfil["volume"]=round(perfil["volume"]*fator_empresa)
    perfil["expedicoes"]=round(perfil["expedicoes"]*fator_empresa)
    perfil["custo"]*=fator_empresa
    perfil["fator"]*=fator_empresa
    perfil["dre"]=round(perfil["custo"]*0.085,2)
    fmt=lambda v: f"{v/1000000:.2f} MM".replace(".",",") if v>=1000000 else f"{v:,.0f}".replace(",",".")
    contexto_cd="Todos os CDs" if cd_selecionado=="Consolidado" else cd_selecionado
    contexto_empresa="Todas as empresas" if empresa_selecionada=="Consolidado" else empresa_selecionada
    contexto=f"{contexto_cd} • {contexto_empresa}"

    kpis=[
        ("📦","VOLUME RECEBIDO",fmt(perfil["volume"]),"Unidades","▲ 8,6% vs mês anterior","#18b84b"),
        ("🚚","EXPEDIÇÕES REALIZADAS",fmt(perfil["expedicoes"]),"Unidades","▲ 7,3% vs mês anterior","#18b84b"),
        ("💲","CUSTO LOGÍSTICO",f"R$ {perfil['custo']:.2f} MM".replace(".",","),"Acumulado no mês","▼ 4,2% vs mês anterior","#e53935"),
        ("🕒","SLA OPERACIONAL",f"{perfil['sla']:.1f}%".replace(".",","),"No mês","▲ 2,1 p.p. vs mês anterior","#18b84b"),
        ("🔄","ÍNDICE DE REVERSA",f"{perfil['reversa']:.2f}%".replace(".",","),"Das expedições","▼ 0,35 p.p. vs mês anterior","#18b84b"),
        ("💰","RESULTADO DO DRE",f"R$ {perfil['dre']:.2f} MM".replace(".",","),"Resultado operacional","▲ 5,8% vs mês anterior","#18b84b"),
    ]
    st.markdown(f'<div class="demo-strip">DADOS DEMONSTRATIVOS • {html.escape(contexto)}</div>',unsafe_allow_html=True)
    for col,item in zip(st.columns(6,gap="small"),kpis):
        with col:
            icone,titulo,valor,unidade,variacao,cor=item
            st.markdown(f'<div class="cio-kpi"><div class="cio-kpi-head"><span class="cio-kpi-icon">{icone}</span><span class="cio-kpi-title">{titulo}</span></div><div class="cio-kpi-value">{valor}</div><div class="cio-kpi-unit">{unidade}</div><div class="cio-kpi-change" style="color:{cor}">{variacao}</div><div class="cio-spark">⌁╱⌁╱╲╱</div></div>',unsafe_allow_html=True)

    cds=pd.DataFrame({
        "CD":["Campinas","Rio de Janeiro","Brasília","Barueri","Contagem","Palhoça","Jaboatão","Manaus","Salvador"],
        "Volume":[238450,205320,141880,186750,152460,98740,88760,69340,98290],
        "SLA":[98.2,96.7,96.1,97.4,95.6,95.2,94.8,93.9,95.0],
        "Reversa":[1.88,2.38,2.42,2.05,2.72,2.81,2.96,3.18,2.89],
        "Ocupação":[84,78,73,88,91,76,82,94,80],
    })
    cds["Volume"]=(cds["Volume"]*fator_empresa).round().astype(int)
    cds["Score"]=(cds["SLA"]-(cds["Reversa"]*1.9)-((cds["Ocupação"]-82).abs()*0.10)).round(1)
    cds["Faixa desempenho"] = cds["Score"].apply(lambda v: "Excelente" if v >= 91 else ("Atenção" if v < 88 else "Estável"))
    cds["Faixa ocupação"] = cds["Ocupação"].apply(lambda v: "Crítica" if v > 90 else ("Adequada" if v >= 75 else "Baixa"))
    cds=cds.sort_values("Score",ascending=False).reset_index(drop=True)
    cds["Rank"]=range(1,len(cds)+1)
    empresas_df=pd.DataFrame({"Empresa":["Claro Móvel","Claro TV","Embratel","NET","Claro Fixo"],"Score":[94.6,90.1,92.7,87.9,89.4]})
    melhor_cd=cds.iloc[0]; pior_cd=cds.iloc[-1]
    melhor_empresa=empresas_df.loc[empresas_df["Score"].idxmax()]; pior_empresa=empresas_df.loc[empresas_df["Score"].idxmin()]

    d1,d2,d3,d4=st.columns(4,gap="small")
    decisoes=[
        (d1,"🏆 MELHOR CD",melhor_cd["CD"],f"Score {melhor_cd['Score']:.1f} • SLA {melhor_cd['SLA']:.1f}%","good"),
        (d2,"⚠️ CD EM ATENÇÃO",pior_cd["CD"],f"Score {pior_cd['Score']:.1f} • Reversa {pior_cd['Reversa']:.2f}%","bad"),
        (d3,"🏆 MELHOR EMPRESA",melhor_empresa["Empresa"],f"Score {melhor_empresa['Score']:.1f}","good"),
        (d4,"⚠️ EMPRESA EM ATENÇÃO",pior_empresa["Empresa"],f"Score {pior_empresa['Score']:.1f}","bad"),
    ]
    for col,titulo,nome,detalhe,classe in decisoes:
        with col:
            st.markdown(f'<div class="decision-card {classe}"><span>{titulo}</span><strong>{html.escape(str(nome))}</strong><small>{html.escape(str(detalhe))}</small></div>',unsafe_allow_html=True)

    alertas_base=[("🔴","SLA abaixo da meta","SLA inferior à meta operacional","Há 2h"),("🟠","Ocupação crítica","CD acima de 90% de ocupação","Há 3h"),("🟠","Pendências de faturamento","Valores aguardando tratamento","Há 4h"),("🔴","Atraso em pedidos Canal Vermelho","Pedidos críticos em atraso","Há 5h")]
    c1,c2,c3=st.columns([1.0,1.18,1.0],gap="medium")
    with c1:
        st.markdown(f'<div class="cio-panel-title">DESEMPENHO DOS CDs</div><div class="cio-panel-sub">Ranking executivo • {html.escape(contexto_empresa)}</div>',unsafe_allow_html=True)
        perf=cds.sort_values("Score")
        perf_chart=alt.Chart(perf).mark_bar(cornerRadiusEnd=5).encode(
            x=alt.X("Score:Q",scale=alt.Scale(domain=[84,94]),axis=alt.Axis(labelColor="#cbd5df",gridColor="#362329",title=None)),
            y=alt.Y("CD:N",sort=alt.EncodingSortField(field="Score",order="descending"),axis=alt.Axis(labelColor="#f2f4f7",title=None)),
            color=alt.Color("Faixa desempenho:N", scale=alt.Scale(domain=["Excelente","Estável","Atenção"], range=["#1dc267","#f2a51a","#ef3340"]), legend=None),
            tooltip=["Rank","CD","Score","SLA","Reversa","Ocupação","Faixa desempenho"]
        ).properties(height=300).configure_view(stroke=None).configure(background="#100c0f")
        st.altair_chart(perf_chart,use_container_width=True)
    with c2:
        st.markdown(f'<div class="cio-panel-title">RANKING DOS CENTROS DE DISTRIBUIÇÃO</div><div class="cio-panel-sub">Verde = melhor desempenho • Vermelho = requer ação</div>',unsafe_allow_html=True)
        df_cd=cds.copy() if cd_selecionado=="Consolidado" else cds[cds["CD"]==cd_selecionado].copy()
        linhas=[]
        for r in df_cd.itertuples():
            status="Excelente" if r.Score>=91 else ("Atenção" if r.Score<88 else "Estável")
            cls="ok" if r.Score>=91 else ("risk" if r.Score<88 else "warn")
            linhas.append(f'<tr class="{cls}"><td><b>#{r.Rank}</b></td><td><strong>{html.escape(r.CD)}</strong></td><td>{r.Volume:,.0f}</td><td>{r.SLA:.1f}%</td><td>{r.Reversa:.2f}%</td><td>{r.Ocupação:.0f}%</td><td><span class="status-pill {cls}">{status}</span></td></tr>'.replace(",","X").replace(".",",").replace("X","."))
        st.markdown(f'<div class="cio-table-wrap enhanced"><table class="cio-table"><thead><tr><th>#</th><th>CD Origem</th><th>Volume</th><th>SLA</th><th>Reversa</th><th>Ocup.</th><th>Status</th></tr></thead><tbody>{"".join(linhas)}</tbody></table></div>',unsafe_allow_html=True)
    with c3:
        st.markdown(f'<div class="cio-panel-title">ALERTAS E DESVIOS CRÍTICOS <span class="alert-count">{len(alertas_base)}</span></div>',unsafe_allow_html=True)
        for icone,titulo,detalhe,tempo in alertas_base:
            st.markdown(f'<div class="cio-alert"><span class="cio-alert-icon">{icone}</span><span><b>{titulo}</b><small>{detalhe} • {html.escape(contexto)}</small></span><em>{tempo}</em></div>',unsafe_allow_html=True)

    meses=["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov"]
    fator=perfil["fator"]
    canal=pd.DataFrame({
        "Mês":meses,
        "Pedidos Sem Canal Vermelho":[round(v*fator) for v in [520,920,2710,410,340,370,420,1130,2740,560,990]],
        "Pedidos Com Canal Vermelho":[round(v*fator) for v in [72,118,286,48,39,43,55,132,305,68,116]],
    }).melt("Mês",var_name="Tipo",value_name="Pedidos")
    ocupacao=cds.sort_values("Ocupação",ascending=False)
    p1,p2,p3=st.columns([1.12,1.0,0.95],gap="medium")
    with p1:
        st.markdown('<div class="cio-panel-title">PEDIDOS DE CANAL VERMELHO</div><div class="cio-panel-sub">Comparativo mensal entre pedidos com e sem Canal Vermelho</div>',unsafe_allow_html=True)
        canal_chart=alt.Chart(canal).mark_bar().encode(
            x=alt.X("Mês:N",sort=meses,axis=alt.Axis(labelColor="#cbd5df",title=None)),
            y=alt.Y("Pedidos:Q",axis=alt.Axis(labelColor="#cbd5df",gridColor="#362329",title=None)),
            color=alt.Color("Tipo:N",scale=alt.Scale(domain=["Pedidos Sem Canal Vermelho","Pedidos Com Canal Vermelho"],range=["#f5a623","#ef3340"]),legend=alt.Legend(labelColor="#fff",title=None)),
            xOffset="Tipo:N",tooltip=["Mês","Tipo","Pedidos"]
        ).properties(height=270).configure_view(stroke=None).configure(background="#100c0f")
        st.altair_chart(canal_chart,use_container_width=True)
    with p2:
        st.markdown('<div class="cio-panel-title">OCUPAÇÃO DOS CDs</div><div class="cio-panel-sub">Faixa recomendada: 75% a 90%</div>',unsafe_allow_html=True)
        ocup_chart=alt.Chart(ocupacao).mark_bar(cornerRadiusEnd=5).encode(
            x=alt.X("Ocupação:Q",scale=alt.Scale(domain=[0,100]),axis=alt.Axis(labelColor="#cbd5df",gridColor="#362329",title=None,format=".0f")),
            y=alt.Y("CD:N",sort="-x",axis=alt.Axis(labelColor="#f2f4f7",title=None)),
            color=alt.Color("Faixa ocupação:N", scale=alt.Scale(domain=["Adequada","Baixa","Crítica"], range=["#1dc267","#f5a623","#ef3340"]), legend=None),
            tooltip=["CD","Ocupação","Faixa ocupação"]
        ).properties(height=270).configure_view(stroke=None).configure(background="#100c0f")
        st.altair_chart(ocup_chart,use_container_width=True)
    with p3:
        st.markdown('<div class="cio-panel-title">DESTAQUES PARA DECISÃO</div>',unsafe_allow_html=True)
        destaques=[
            f"Melhor CD: {melhor_cd['CD']} (score {melhor_cd['Score']:.1f})",
            f"CD em atenção: {pior_cd['CD']} (ocupação {pior_cd['Ocupação']:.0f}%)",
            f"Melhor empresa: {melhor_empresa['Empresa']}",
            f"Empresa em atenção: {pior_empresa['Empresa']}",
            f"Filtro atual: {contexto}",
        ]
        for i,texto in enumerate(destaques):
            ic="🏆" if i in (0,2) else ("⚠️" if i in (1,3) else "ℹ️")
            st.markdown(f'<div class="cio-highlight">{ic} {html.escape(str(texto))}</div>',unsafe_allow_html=True)


def exibir_atualizacoes():
    st.markdown('<div class="overview-title">Atualizações dos indicadores</div>',unsafe_allow_html=True)
    st.caption("Periodicidade, referência dos dados e disponibilidade dos painéis.")
    df=pd.DataFrame(catalogo_visao_geral()); area=st.selectbox("Área",["Todas"]+list(INDICADORES.keys()),key="filtro_area_atualizacoes")
    if area!="Todas": df=df[df["Área"]==area]
    st.dataframe(df,use_container_width=True,hide_index=True)


def exibir_minha_conta():
    u=st.session_state.usuario_logado or {}
    st.markdown('<div class="overview-title">Minha conta</div>',unsafe_allow_html=True)
    with st.container(border=True):
        c1,c2=st.columns(2)
        with c1: st.write(f"**Nome:** {u.get('nome_completo') or 'Não informado'}"); st.write(f"**Usuário:** {u.get('usuario') or 'Não informado'}"); st.write(f"**E-mail:** {u.get('email') or 'Não informado'}")
        with c2: st.write(f"**Perfil:** {u.get('perfil') or 'USUARIO'}"); st.write(f"**Aprovado em:** {data_hora_brasilia(u.get('aprovado_em'))}"); st.write("**Status:** Ativo")


def exibir_portal_usuario():
    u=st.session_state.usuario_logado or {}; nome=u.get("nome_completo") or u.get("usuario") or "Usuário"
    with st.sidebar:
        st.markdown('<div class="side-brand">Central Operacional</div>',unsafe_allow_html=True); st.caption(f"Acesso de {nome}")
        pagina=st.radio("Navegação",["🏠 Visão Geral","📊 Indicadores","🗄️ Fontes de dados","📅 Atualizações","👤 Minha conta"],label_visibility="collapsed",key="pagina_portal_usuario")
        st.divider()
        if st.button("Sair",use_container_width=True,key="sair_menu_lateral"):
            st.session_state.usuario_logado=None; st.session_state.pop("pagina_portal_usuario",None); st.rerun()
    if pagina=="🏠 Visão Geral": exibir_visao_geral()
    elif pagina=="📊 Indicadores": exibir_central_indicadores()
    elif pagina=="🗄️ Fontes de dados": exibir_fontes_dados()
    elif pagina=="📅 Atualizações": exibir_atualizacoes()
    else: exibir_minha_conta()


def exibir_central_indicadores():
    usuario_atual = st.session_state.usuario_logado or {}
    st.markdown('<div class="overview-title">Indicadores operacionais</div>', unsafe_allow_html=True)
    st.caption("Selecione a área e clique no card para abrir o painel detalhado.")
    area = st.segmented_control("Área operacional", list(INDICADORES.keys()), default="Armazenagem", key="area_indicadores_padrao", label_visibility="collapsed") or "Armazenagem"
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
                periodicidade = indicador.get("periodicidade") or "A definir"
                ultima_atualizacao = indicador.get("ultima_atualizacao") or "-"
                rodape_atualizacao = f"{periodicidade} • Atualizado até {ultima_atualizacao}"
                st.markdown(
                    f"""<div class="card-indicador-visual">
                        <div class="card-indicador-principal">
                            <span class="card-indicador-icone">{html.escape(str(icone))}</span>
                            <span class="card-indicador-titulo">{html.escape(str(titulo))}</span>
                        </div>
                        <div class="card-indicador-rodape">{html.escape(rodape_atualizacao)}</div>
                    </div>""",
                    unsafe_allow_html=True,
                )
                if st.button(
                    f"Abrir {titulo}",
                    key=f"card_indicador_{area}_{inicio}_{titulo}",
                    use_container_width=True,
                    type="secondary",
                ):
                    registrar_acesso_indicador(usuario_atual, area, titulo, destino)
                    st.session_state.indicador_destino_pendente = destino
                    st.rerun()

    destino_pendente = st.session_state.pop("indicador_destino_pendente", None)
    if destino_pendente:
        destino_js = json.dumps(destino_pendente)
        components_html(
            f"<script>window.parent.open({destino_js}, '_blank', 'noopener,noreferrer');</script>",
            height=0,
        )


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
        CREATE TABLE IF NOT EXISTS acessos_indicadores (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario_id INTEGER,
            nome_completo TEXT,
            usuario TEXT,
            area TEXT NOT NULL,
            indicador TEXT NOT NULL,
            url TEXT NOT NULL,
            acessado_em TEXT NOT NULL,
            sessao_id TEXT,
            FOREIGN KEY(usuario_id) REFERENCES usuarios(id)
        );
        CREATE INDEX IF NOT EXISTS idx_acessos_indicadores_data
            ON acessos_indicadores(acessado_em);
        CREATE INDEX IF NOT EXISTS idx_acessos_indicadores_usuario
            ON acessos_indicadores(usuario_id);
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
    with conectar() as con:
        existe = con.execute("SELECT 1 FROM usuarios WHERE nome_hash=? OR email_hash=?", (nome_hash, email_hash)).fetchone()
        pendente = con.execute("SELECT 1 FROM solicitacoes WHERE status='PENDENTE' AND (nome_hash=? OR email_hash=?)", (nome_hash, email_hash)).fetchone()
        if existe or pendente:
            raise ValueError("Já existe cadastro ou solicitação pendente para os dados informados.")
        con.execute("""INSERT INTO solicitacoes
            (tipo,nome_completo,usuario,email,nome_hash,email_hash,senha_hash_pendente,salt_pendente,iteracoes,codigo_dispositivo_hash,status,criado_em)
            VALUES('CADASTRO',?,?,?,?,?,?,?,?,?,'PENDENTE',?)""",
            (nome.strip(), usuario, email, nome_hash, email_hash, senha_hash, salt, ITERACOES, None, agora_iso()))
    return True


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


def gerar_url_email_status_solicitacao(solicitacao, aprovado, observacao=""):
    nome = solicitacao.get("nome_completo") or solicitacao.get("usuario") or "Usuário"
    destinatario = solicitacao.get("email") or ""
    tipo = solicitacao.get("tipo")
    observacao = str(observacao or "").strip()

    if tipo == "CADASTRO" and aprovado:
        assunto = "Central de Inteligência Operacional | Cadastro aprovado"
        corpo = f"""Olá, {nome}.

Informamos que sua solicitação de acesso à Central de Inteligência Operacional foi aprovada.

Dados do cadastro:
Usuário: {solicitacao.get('usuario') or '-'}
E-mail: {destinatario}
Status: Aprovado

Você já pode acessar a Central utilizando seu usuário, e-mail corporativo e a senha cadastrada.

Em caso de dificuldade no acesso, entre em contato com a equipe responsável.

Central de Inteligência Operacional"""
    elif tipo == "CADASTRO":
        assunto = "Central de Inteligência Operacional | Solicitação de cadastro"
        corpo = f"""Olá, {nome}.

Informamos que sua solicitação de acesso à Central de Inteligência Operacional não foi aprovada neste momento.

Status da solicitação: Reprovada
Observação: {observacao or 'Não informada'}

Caso necessite de esclarecimentos ou de uma nova análise, entre em contato com a equipe responsável.

Central de Inteligência Operacional"""
    elif aprovado:
        assunto = "Central de Inteligência Operacional | Redefinição de senha aprovada"
        corpo = f"""Olá, {nome}.

Sua solicitação de redefinição de senha da Central de Inteligência Operacional foi aprovada.

A nova senha cadastrada na solicitação já está ativa e poderá ser utilizada no próximo acesso.

Por segurança, nunca compartilhe sua senha com outras pessoas.

Central de Inteligência Operacional"""
    else:
        assunto = "Central de Inteligência Operacional | Solicitação de redefinição de senha"
        corpo = f"""Olá, {nome}.

Informamos que sua solicitação de redefinição de senha não foi aprovada neste momento.

Status da solicitação: Reprovada
Observação: {observacao or 'Não informada'}

Caso não reconheça essa solicitação ou precise de uma nova análise, entre em contato com a equipe responsável.

Central de Inteligência Operacional"""

    return (
        "https://outlook.office.com/mail/deeplink/compose"
        f"?to={quote(destinatario)}&subject={quote(assunto)}&body={quote(corpo)}"
    )


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


def gerar_relacao_usuarios_xlsx(usuarios):
    colunas = [
        "Nome",
        "Usuário",
        "E-mail",
        "Perfil",
        "Status atual",
        "Criado em",
        "Aprovado em",
        "Atualizado em",
    ]

    wb = Workbook()
    ws = wb.active
    ws.title = "Usuarios"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A2"

    preenchimento_cabecalho = PatternFill("solid", fgColor="C00000")
    fonte_cabecalho = Font(color="FFFFFF", bold=True, size=11)
    borda_inferior = Border(bottom=Side(style="thin", color="D9D9D9"))

    for coluna, titulo in enumerate(colunas, start=1):
        celula = ws.cell(row=1, column=coluna, value=titulo)
        celula.fill = preenchimento_cabecalho
        celula.font = fonte_cabecalho
        celula.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 32

    for linha, item in enumerate(usuarios, start=2):
        status_atual = "Ativo" if item.get("ativo") else "Desativado"
        valores = [
            item.get("nome_completo") or "Cadastro legado",
            str(item.get("usuario") or "Não disponível"),
            item.get("email") or "Não disponível",
            str(item.get("perfil") or "USUARIO").upper(),
            status_atual,
            data_hora_brasilia(item.get("criado_em")) if item.get("criado_em") else "-",
            data_hora_brasilia(item.get("aprovado_em")) if item.get("aprovado_em") else "-",
            data_hora_brasilia(item.get("atualizado_em")) if item.get("atualizado_em") else "-",
        ]
        for coluna, valor in enumerate(valores, start=1):
            celula = ws.cell(row=linha, column=coluna, value=valor)
            celula.alignment = Alignment(vertical="center", wrap_text=True)
            celula.border = borda_inferior
            if coluna in (2, 3):
                celula.number_format = "@"
        ws.row_dimensions[linha].height = 28

    larguras = {
        "A": 35,
        "B": 24,
        "C": 36,
        "D": 18,
        "E": 18,
        "F": 24,
        "G": 24,
        "H": 24,
    }
    for coluna, largura in larguras.items():
        ws.column_dimensions[coluna].width = largura

    ws.auto_filter.ref = f"A1:H{max(ws.max_row, 1)}"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_title_rows = "1:1"

    arquivo = BytesIO()
    wb.save(arquivo)
    arquivo.seek(0)
    return arquivo.getvalue()


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
    incluidos = atualizados = ignorados = 0
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
                iteracoes = int(item.get("iteracoes", ITERACOES))
                bytes.fromhex(str(item["salt"]))
            except (TypeError, ValueError):
                ignorados += 1
                erros.append(f"Registro {indice}: salt ou iterações inválidos.")
                continue

            existente = con.execute(
                """SELECT id,nome_completo,usuario,email,perfil,criado_em,aprovado_em,
                          atualizado_em,codigo_dispositivo_hash
                   FROM usuarios WHERE nome_hash=? OR email_hash=?""",
                (item["nome_hash"], item["email_hash"]),
            ).fetchone()

            agora = agora_iso()
            criado_em = item.get("criado_em") or (existente["criado_em"] if existente else None) or item.get("aprovado_em") or agora
            aprovado_em = item.get("aprovado_em") or (existente["aprovado_em"] if existente else None)
            atualizado_em = item.get("atualizado_em") or agora
            perfil = str(item.get("perfil") or (existente["perfil"] if existente else None) or "USUARIO")
            ativo = 1 if item.get("ativo", True) else 0
            valores = (
                item.get("nome_completo") or (existente["nome_completo"] if existente else None),
                item.get("usuario") or (existente["usuario"] if existente else None),
                item.get("email") or (existente["email"] if existente else None),
                item["nome_hash"],
                item["email_hash"],
                item["senha_hash"],
                item["salt"],
                iteracoes,
                item.get("codigo_dispositivo_hash") or (existente["codigo_dispositivo_hash"] if existente else None),
                ativo,
                perfil,
                criado_em,
                aprovado_em,
                atualizado_em,
            )

            if existente:
                con.execute(
                    """UPDATE usuarios SET
                       nome_completo=?,usuario=?,email=?,nome_hash=?,email_hash=?,senha_hash=?,salt=?,iteracoes=?,
                       codigo_dispositivo_hash=?,ativo=?,perfil=?,criado_em=?,aprovado_em=?,atualizado_em=?
                       WHERE id=?""",
                    valores + (existente["id"],),
                )
                atualizados += 1
            else:
                con.execute(
                    """INSERT INTO usuarios
                       (nome_completo,usuario,email,nome_hash,email_hash,senha_hash,salt,iteracoes,
                        codigo_dispositivo_hash,ativo,perfil,criado_em,aprovado_em,atualizado_em)
                       VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    valores,
                )
                incluidos += 1

        acessos_backup = dados.get("acessos_indicadores", []) if isinstance(dados, dict) else []
        if isinstance(acessos_backup, list):
            for acesso in acessos_backup:
                if not isinstance(acesso, dict) or not acesso.get("indicador") or not acesso.get("acessado_em"):
                    continue
                duplicado_acesso = con.execute(
                    """SELECT 1 FROM acessos_indicadores
                       WHERE COALESCE(usuario_id,-1)=COALESCE(?,-1) AND indicador=? AND acessado_em=? AND COALESCE(sessao_id,'')=COALESCE(?,'')""",
                    (acesso.get("usuario_id"), acesso.get("indicador"), acesso.get("acessado_em"), acesso.get("sessao_id")),
                ).fetchone()
                if not duplicado_acesso:
                    con.execute(
                        """INSERT INTO acessos_indicadores
                           (usuario_id,nome_completo,usuario,area,indicador,url,acessado_em,sessao_id)
                           VALUES(?,?,?,?,?,?,?,?)""",
                        (
                            acesso.get("usuario_id"), acesso.get("nome_completo"), acesso.get("usuario"),
                            acesso.get("area") or "Não informada", acesso.get("indicador"), acesso.get("url") or "",
                            acesso.get("acessado_em"), acesso.get("sessao_id"),
                        ),
                    )

        con.execute(
            "INSERT INTO auditoria(acao,referencia,administrador,criado_em) VALUES(?,?,?,?)",
            (
                "RESTAURAR_BACKUP_USUARIOS_JSON",
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


def exportar_backup_completo_usuarios():
    campos = (
        "nome_completo,usuario,email,nome_hash,email_hash,senha_hash,salt,iteracoes,"
        "codigo_dispositivo_hash,ativo,perfil,criado_em,aprovado_em,atualizado_em"
    )
    with conectar() as con:
        linhas = con.execute(
            f"SELECT {campos} FROM usuarios ORDER BY id"
        ).fetchall()
    return json.dumps(
        {
            "tipo": "BACKUP_COMPLETO_USUARIOS_CENTRAL",
            "versao": 29,
            "usuarios": [dict(linha) for linha in linhas],
            "acessos_indicadores": [dict(linha) for linha in conectar().execute(
                "SELECT usuario_id,nome_completo,usuario,area,indicador,url,acessado_em,sessao_id FROM acessos_indicadores ORDER BY id"
            ).fetchall()],
            "atualizado_em": agora_iso(),
        },
        ensure_ascii=False,
        indent=2,
    )


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

/* Cards funcionais com registro de acesso */
.indicador-card-conteudo{position:relative;display:flex;min-height:150px;padding:22px 16px 16px;border:1px solid rgba(218,41,28,.14);border-radius:22px 22px 8px 8px;background:linear-gradient(145deg,#fff,#fff6f3);box-shadow:0 8px 20px rgba(79,20,20,.09);flex-direction:column;align-items:center;justify-content:center;text-align:center}
.indicador-card-conteudo:before{content:"";position:absolute;left:0;top:0;width:100%;height:5px;background:linear-gradient(90deg,#b51f25,#da291c 50%,#f47b45)}
.indicador-card-conteudo + div[data-testid="stButton"] button{border-radius:8px 8px 14px 14px!important;border:1px solid rgba(218,41,28,.25)!important;background:#fff!important;color:#b51f25!important;font-weight:800!important}
.indicador-card-conteudo + div[data-testid="stButton"] button:hover{background:#fff1f1!important;border-color:#da291c!important}
@media(max-width:580px){.indicador-card-conteudo{min-height:120px;padding:16px 12px 12px;border-radius:17px 17px 8px 8px}}

/* Card funcional único, sem botão Abrir separado */
[class*="st-key-card_indicador_"] button{min-height:216px!important;width:100%!important;padding:24px 16px!important;border:1px solid rgba(218,41,28,.16)!important;border-top:5px solid #ed5b35!important;border-radius:22px!important;background:linear-gradient(145deg,#fff,#fff6f3)!important;box-shadow:0 10px 28px rgba(79,20,20,.10)!important;color:#202231!important;white-space:pre-line!important;transition:transform .2s ease,box-shadow .2s ease!important}
[class*="st-key-card_indicador_"] button p{font-size:1rem!important;line-height:1.55!important;font-weight:850!important;white-space:pre-line!important;text-align:center!important}
[class*="st-key-card_indicador_"] button p:first-line{font-size:2rem!important;line-height:2!important}
[class*="st-key-card_indicador_"] button:hover{transform:translateY(-5px)!important;border-color:#da291c!important;box-shadow:0 18px 36px rgba(122,25,25,.18)!important;background:linear-gradient(145deg,#fff,#fff1ed)!important}
@media(max-width:580px){[class*="st-key-card_indicador_"] button{min-height:150px!important;padding:18px 10px!important;border-radius:17px!important}[class*="st-key-card_indicador_"] button p{font-size:.88rem!important}}

/* Cards mais estreitos, compactos e centralizados */
div[data-testid="stHorizontalBlock"]:has([class*="st-key-card_indicador_"]){
  width:min(100%,1480px)!important;
  max-width:1480px!important;
  margin-left:auto!important;
  margin-right:auto!important;
  gap:18px!important;
  justify-content:center!important;
}
div[data-testid="stHorizontalBlock"]:has([class*="st-key-card_indicador_"])>div{
  min-width:0!important;
}
[class*="st-key-card_indicador_"] button{
  min-height:174px!important;
  height:174px!important;
  padding:18px 12px!important;
  border-radius:18px!important;
}
[class*="st-key-card_indicador_"] button p{
  font-size:.94rem!important;
  line-height:1.35!important;
}
[class*="st-key-card_indicador_"] button p:first-line{
  font-size:1.75rem!important;
  line-height:1.65!important;
}
@media(max-width:1200px){
  div[data-testid="stHorizontalBlock"]:has([class*="st-key-card_indicador_"]){
    width:100%!important;
    gap:12px!important;
  }
  [class*="st-key-card_indicador_"] button{
    min-height:164px!important;
    height:164px!important;
    padding:16px 10px!important;
  }
  [class*="st-key-card_indicador_"] button p{font-size:.86rem!important}
}
@media(max-width:900px){
  div[data-testid="stHorizontalBlock"]:has([class*="st-key-card_indicador_"]){
    display:grid!important;
    grid-template-columns:repeat(2,minmax(0,1fr))!important;
    gap:12px!important;
  }
  div[data-testid="stHorizontalBlock"]:has([class*="st-key-card_indicador_"])>div{
    width:100%!important;
    min-width:0!important;
    max-width:none!important;
    flex:none!important;
  }
  [class*="st-key-card_indicador_"] button{
    min-height:150px!important;
    height:150px!important;
  }
}
@media(max-width:580px){
  div[data-testid="stHorizontalBlock"]:has([class*="st-key-card_indicador_"]){
    grid-template-columns:1fr!important;
    gap:10px!important;
  }
  [class*="st-key-card_indicador_"] button{
    min-height:126px!important;
    height:126px!important;
    padding:14px 10px!important;
    border-radius:15px!important;
  }
  [class*="st-key-card_indicador_"] button p{
    font-size:.86rem!important;
    line-height:1.25!important;
  }
}

/* Informações discretas de periodicidade e última atualização */
[class*="st-key-card_indicador_"] button{
  min-height:190px!important;
  height:190px!important;
}
[class*="st-key-card_indicador_"] button p{
  white-space:pre-line!important;
  line-height:1.32!important;
}
[class*="st-key-card_indicador_"] button p::first-line{
  font-size:1.72rem!important;
}
@media(max-width:1200px){
  [class*="st-key-card_indicador_"] button{min-height:180px!important;height:180px!important}
}
@media(max-width:900px){
  [class*="st-key-card_indicador_"] button{min-height:166px!important;height:166px!important}
}
@media(max-width:580px){
  [class*="st-key-card_indicador_"] button{min-height:142px!important;height:142px!important}
}

/* Card com título destacado e metadados discretos no rodapé */
.card-indicador-visual{
  position:relative;
  display:flex;
  flex-direction:column;
  justify-content:center;
  align-items:center;
  width:100%;
  height:190px;
  box-sizing:border-box;
  padding:22px 16px 34px;
  overflow:hidden;
  border:1px solid rgba(218,41,28,.16);
  border-top:5px solid #ed5b35;
  border-radius:18px;
  background:linear-gradient(145deg,#fff,#fff6f3);
  box-shadow:0 10px 28px rgba(79,20,20,.10);
  pointer-events:none;
  transition:transform .2s ease,box-shadow .2s ease,border-color .2s ease;
}
.card-indicador-principal{
  display:flex;
  align-items:center;
  justify-content:center;
  gap:7px;
  width:100%;
  text-align:center;
}
.card-indicador-icone{font-size:1.12rem;line-height:1}
.card-indicador-titulo{
  color:#202231;
  font-size:.94rem;
  font-weight:850;
  line-height:1.28;
}
.card-indicador-rodape{
  position:absolute;
  left:10px;
  right:10px;
  bottom:13px;
  color:#777b85;
  font-size:.67rem;
  font-weight:400;
  line-height:1.15;
  text-align:center;
  white-space:normal;
}
/* O botão Streamlit fica transparente sobre o card inteiro. */
[class*="st-key-card_indicador_"]{
  position:relative!important;
  z-index:3!important;
  height:190px!important;
  margin-top:-190px!important;
}
[class*="st-key-card_indicador_"] button{
  width:100%!important;
  min-height:190px!important;
  height:190px!important;
  padding:0!important;
  border:0!important;
  border-radius:18px!important;
  background:transparent!important;
  box-shadow:none!important;
  color:transparent!important;
  opacity:1!important;
}
[class*="st-key-card_indicador_"] button p{font-size:0!important;color:transparent!important}
div[data-testid="stVerticalBlock"]:has(> .card-indicador-visual):hover .card-indicador-visual{
  transform:translateY(-4px);
  border-color:#da291c;
  box-shadow:0 16px 34px rgba(122,25,25,.16);
}
@media(max-width:1200px){
  .card-indicador-visual,[class*="st-key-card_indicador_"],[class*="st-key-card_indicador_"] button{
    height:180px!important;min-height:180px!important;
  }
  [class*="st-key-card_indicador_"]{margin-top:-180px!important}
}
@media(max-width:900px){
  .card-indicador-visual,[class*="st-key-card_indicador_"],[class*="st-key-card_indicador_"] button{
    height:166px!important;min-height:166px!important;
  }
  [class*="st-key-card_indicador_"]{margin-top:-166px!important}
  .card-indicador-rodape{font-size:.65rem;bottom:11px}
}
@media(max-width:580px){
  .card-indicador-visual,[class*="st-key-card_indicador_"],[class*="st-key-card_indicador_"] button{
    height:142px!important;min-height:142px!important;
  }
  .card-indicador-visual{padding:16px 10px 30px;border-radius:15px}
  [class*="st-key-card_indicador_"]{height:142px!important;margin-top:-142px!important}
  [class*="st-key-card_indicador_"] button{border-radius:15px!important}
  .card-indicador-titulo{font-size:.86rem}
  .card-indicador-rodape{font-size:.62rem;bottom:9px}
}

/* Refinamento final da tipografia dos cards */
.card-indicador-titulo{
  font-size:1.01rem!important;
  line-height:1.30!important;
  font-weight:850!important;
}
.card-indicador-rodape{
  color:#565b65!important;
  font-size:.73rem!important;
  font-weight:400!important;
  line-height:1.18!important;
  bottom:13px!important;
}
@media(max-width:1200px){
  .card-indicador-titulo{font-size:.97rem!important}
  .card-indicador-rodape{font-size:.71rem!important}
}
@media(max-width:900px){
  .card-indicador-titulo{font-size:.93rem!important}
  .card-indicador-rodape{font-size:.69rem!important;bottom:11px!important}
}
@media(max-width:580px){
  .card-indicador-titulo{font-size:.89rem!important}
  .card-indicador-rodape{font-size:.66rem!important;bottom:9px!important}
}

/* Mantem exatamente as mesmas cores do card durante hover, foco e clique */
.card-indicador-visual:hover,
div[data-testid="stVerticalBlock"]:has(> .card-indicador-visual):hover .card-indicador-visual,
div[data-testid="stVerticalBlock"]:has(> .card-indicador-visual):focus-within .card-indicador-visual{
  background:linear-gradient(145deg,#fff,#fff6f3)!important;
  border-color:rgba(218,41,28,.16)!important;
  border-top-color:#ed5b35!important;
  box-shadow:0 10px 28px rgba(79,20,20,.10)!important;
  transform:none!important;
  filter:none!important;
}
[class*="st-key-card_indicador_"] button,
[class*="st-key-card_indicador_"] button:hover,
[class*="st-key-card_indicador_"] button:focus,
[class*="st-key-card_indicador_"] button:focus-visible,
[class*="st-key-card_indicador_"] button:active{
  background:transparent!important;
  background-color:transparent!important;
  border:0!important;
  box-shadow:none!important;
  color:transparent!important;
  opacity:1!important;
  transform:none!important;
  filter:none!important;
  outline:none!important;
}
[class*="st-key-card_indicador_"] button:hover *,
[class*="st-key-card_indicador_"] button:focus *,
[class*="st-key-card_indicador_"] button:active *{
  background:transparent!important;
  color:transparent!important;
}

/* Hover: desloca o card para cima sem alterar fundo, bordas ou cores */
.card-indicador-visual:hover,
div[data-testid="stVerticalBlock"]:has(> .card-indicador-visual):hover .card-indicador-visual,
div[data-testid="stVerticalBlock"]:has(> .card-indicador-visual):focus-within .card-indicador-visual{
  background:linear-gradient(145deg,#fff,#fff6f3)!important;
  border-color:rgba(218,41,28,.16)!important;
  border-top-color:#ed5b35!important;
  box-shadow:0 16px 34px rgba(122,25,25,.16)!important;
  transform:translateY(-4px)!important;
  filter:none!important;
}
[class*="st-key-card_indicador_"] button:hover,
[class*="st-key-card_indicador_"] button:focus,
[class*="st-key-card_indicador_"] button:focus-visible,
[class*="st-key-card_indicador_"] button:active{
  background:transparent!important;
  background-color:transparent!important;
  border:0!important;
  box-shadow:none!important;
  color:transparent!important;
  opacity:1!important;
  transform:none!important;
  filter:none!important;
  outline:none!important;
}

/* Correcao definitiva do hover: aplica o movimento na coluna que contem o botao transparente */
div[data-testid="stColumn"]:has([class*="st-key-card_indicador_"] button:hover) .card-indicador-visual,
div[data-testid="column"]:has([class*="st-key-card_indicador_"] button:hover) .card-indicador-visual,
div[data-testid="stColumn"]:has([class*="st-key-card_indicador_"]:hover) .card-indicador-visual,
div[data-testid="column"]:has([class*="st-key-card_indicador_"]:hover) .card-indicador-visual{
  background:linear-gradient(145deg,#fff,#fff6f3)!important;
  border-color:rgba(218,41,28,.16)!important;
  border-top-color:#ed5b35!important;
  box-shadow:0 16px 34px rgba(122,25,25,.16)!important;
  transform:translateY(-4px)!important;
  filter:none!important;
}
/* A camada clicavel permanece totalmente transparente em todos os estados. */
[class*="st-key-card_indicador_"] button:hover,
[class*="st-key-card_indicador_"] button:focus,
[class*="st-key-card_indicador_"] button:focus-visible,
[class*="st-key-card_indicador_"] button:active{
  background:transparent!important;
  background-color:transparent!important;
  border-color:transparent!important;
  box-shadow:none!important;
  color:transparent!important;
  opacity:1!important;
  filter:none!important;
  outline:none!important;
}

[data-testid="stSidebar"]{background:linear-gradient(180deg,#fff,#f7f8fa);border-right:1px solid #e4e7eb}
[data-testid="stSidebar"] [data-testid="stRadio"] label{padding:.55rem .65rem;border-radius:10px}
[data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked){background:#fff0ee;color:#c5221f;font-weight:750}
.side-brand{font-size:1.15rem;font-weight:850;color:#b51f25}.overview-title{font-size:1.58rem;font-weight:850;color:#202633;margin:.1rem 0 .25rem}
.overview-kpi{min-height:118px;padding:16px 15px;border-radius:16px;background:#fff;border:1px solid #e5e7eb;border-top:4px solid #c9cdd3;box-shadow:0 7px 20px rgba(28,36,48,.07)}
.overview-kpi.positivo{border-top-color:#35a853}.overview-kpi.alerta{border-top-color:#ef8b31}.overview-kpi.destaque{border-top-color:#d52b1e;background:linear-gradient(145deg,#fff,#fff5f2)}
.overview-kpi-label{font-size:.72rem;color:#69707d;text-transform:uppercase;font-weight:750}.overview-kpi-value{font-size:1.45rem;color:#202633;font-weight:850;line-height:1.15;margin:8px 0 4px;overflow-wrap:anywhere}.overview-kpi-detail{font-size:.72rem;color:#707784}
.quick-card{height:122px;padding:15px;border-radius:15px;background:#fff;border:1px solid #e4e6e9;border-top:4px solid #ed5b35;box-shadow:0 7px 18px rgba(35,42,55,.07)}.quick-icon{font-size:1.25rem}.quick-name{font-size:.88rem;font-weight:820;color:#202633;margin:8px 0 7px}.quick-meta{font-size:.68rem;color:#6b7280}

/* Painel executivo escuro inspirado no modelo de referência */
.general-dashboard-title{font-size:1.7rem;color:#202633;font-weight:850;line-height:1.15;margin:.2rem 0}.general-dashboard-subtitle{font-size:.78rem;color:#697481}.demo-badge{margin-top:8px;padding:8px 10px;border-radius:8px;background:#8c10d4;color:#fff;text-align:center;font-size:.66rem;font-weight:850;letter-spacing:.4px}
.general-section-title{font-size:1.03rem;color:#202633;font-weight:820;margin:18px 0 8px;padding-left:9px;border-left:4px solid #e53935}.summary-indicator-card{min-height:148px;border:1px solid #7f8a95;border-top:4px solid var(--accent);border-radius:14px;background:#1a2530;padding:13px 12px;box-shadow:0 5px 14px rgba(0,0,0,.22);box-sizing:border-box}.summary-indicator-name{font-size:.72rem;color:#e9eef3;font-weight:760;line-height:1.2;min-height:34px}.summary-indicator-value{font-size:1.55rem;color:#fff;font-weight:900;margin:7px 0 3px;line-height:1}.summary-indicator-label{font-size:.68rem;color:#c9d2db;font-weight:600}.summary-indicator-detail{font-size:.61rem;color:#8fa0ae;margin-top:9px;line-height:1.2}.chart-panel-title{font-size:.86rem;color:#f2f5f8;font-weight:800;background:#19232e;border:1px solid #65717d;border-bottom:0;border-radius:12px 12px 0 0;padding:9px 11px;margin-top:2px}[data-testid="stVegaLiteChart"]{background:#19232e;border:1px solid #65717d;border-top:0;border-radius:0 0 12px 12px;padding:3px 7px 7px;margin-bottom:10px}@media(max-width:900px){.summary-indicator-card{min-height:132px}.summary-indicator-value{font-size:1.3rem}.general-dashboard-title{font-size:1.4rem}}

/* Visão geral executiva no modelo vermelho/preto */
.cio-title{font-size:1.55rem;font-weight:900;color:#fff;line-height:1.1}.cio-subtitle{font-size:.75rem;color:#a9a3a5;margin-top:4px}.cio-update{font-size:.67rem;color:#aaa;padding:7px 0}.cio-update strong{color:#fff;font-size:.78rem}.demo-strip{background:linear-gradient(90deg,#7d0710,#d21722,#7d0710);color:#fff;text-align:center;font-size:.62rem;font-weight:850;letter-spacing:.7px;padding:5px;border-radius:6px;margin:7px 0 11px}.cio-kpi{height:156px;background:linear-gradient(145deg,#1a1013,#0d0b0d);border:1px solid #50272c;border-radius:14px;padding:13px;box-shadow:inset 0 0 20px rgba(180,0,12,.08),0 7px 18px rgba(0,0,0,.3);position:relative}.cio-kpi-head{display:flex;align-items:center;gap:8px}.cio-kpi-icon{display:flex;align-items:center;justify-content:center;width:42px;height:42px;border-radius:50%;background:linear-gradient(145deg,#ef2630,#a70710);font-size:1.3rem}.cio-kpi-title{font-size:.62rem;color:#f4f0f1;font-weight:850;line-height:1.2}.cio-kpi-value{font-size:1.35rem;color:#fff;font-weight:900;margin:10px 0 2px}.cio-kpi-unit{font-size:.62rem;color:#aaa}.cio-kpi-change{font-size:.58rem;font-weight:750;position:absolute;bottom:11px;left:13px}.cio-spark{position:absolute;right:10px;bottom:8px;color:#d7202c;font-size:1.15rem}.cio-panel-title{background:#120d0f;border:1px solid #50272c;border-bottom:0;border-radius:13px 13px 0 0;color:#fff;font-size:.76rem;font-weight:900;padding:10px 12px}.cio-panel-sub{background:#120d0f;color:#999;font-size:.6rem;padding:0 12px 7px;border-left:1px solid #50272c;border-right:1px solid #50272c}.alert-count{background:#c1121f;border-radius:10px;padding:2px 7px;margin-left:5px}.cio-alert{display:grid;grid-template-columns:28px 1fr 42px;gap:7px;align-items:center;background:#100c0e;border:1px solid #3a2428;border-top:0;padding:9px;color:#eee;font-size:.66rem}.cio-alert small{display:block;color:#a6a0a2;margin-top:3px}.cio-alert em{font-style:normal;color:#d0444b;font-size:.58rem;text-align:right}.project-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:5px;background:#100c0e;border:1px solid #50272c;padding:9px}.project-grid div{background:#171013;border:1px solid #41282b;padding:8px;border-radius:8px}.project-grid b,.project-grid small{display:block;color:#bbb;font-size:.58rem}.project-grid strong{display:block;color:#fff;font-size:1.25rem;margin:6px 0}.financial-row{display:grid;grid-template-columns:repeat(3,1fr);background:#100c0e;border:1px solid #50272c;border-top:0;padding:10px}.financial-row div{border-right:1px solid #3c292c;text-align:center}.financial-row div:last-child{border:0}.financial-row small{display:block;color:#aaa;font-size:.55rem}.financial-row b{display:block;color:#fff;font-size:.84rem;margin-top:5px}.financial-row div:first-child b{color:#26c45a}.cio-highlight{background:#100c0e;border:1px solid #3a2428;border-top:0;color:#ddd;padding:10px;font-size:.66rem}[data-testid="stAppViewContainer"]{background:radial-gradient(circle at 20% 0%,#37080d 0,#12090c 28%,#070708 70%)!important}[data-testid="stVegaLiteChart"]{background:#100c0f!important;border-color:#50272c!important}@media(max-width:900px){.cio-kpi{height:140px}.cio-kpi-value{font-size:1.12rem}.project-grid{grid-template-columns:repeat(2,1fr)}}

/* CORRECAO GLOBAL DE LEGIBILIDADE E TEMA HIBRIDO */
:root{
  --cio-page:#f3f5f7;
  --cio-surface:#ffffff;
  --cio-text:#1f2937;
  --cio-muted:#5f6b78;
  --cio-border:#d8dde3;
  --cio-red:#c62026;
  --cio-dark:#100c0f;
}
/* O fundo escuro fica somente na Visao Geral; as demais paginas voltam ao tema claro. */
[data-testid="stAppViewContainer"]{
  background:var(--cio-page)!important;
  color:var(--cio-text)!important;
}
[data-testid="stMainBlockContainer"],.block-container{color:var(--cio-text)!important}
/* Textos gerais e titulos */
[data-testid="stAppViewContainer"] h1,
[data-testid="stAppViewContainer"] h2,
[data-testid="stAppViewContainer"] h3,
[data-testid="stAppViewContainer"] h4,
[data-testid="stAppViewContainer"] p,
[data-testid="stAppViewContainer"] label,
[data-testid="stAppViewContainer"] .stCaption,
[data-testid="stAppViewContainer"] [data-testid="stMarkdownContainer"]{
  color:var(--cio-text);
}
.overview-title{color:#202633!important}
/* Formularios e containers claros */
[data-testid="stForm"],
[data-testid="stExpander"],
[data-testid="stVerticalBlockBorderWrapper"]{
  background:var(--cio-surface)!important;
  color:var(--cio-text)!important;
  border-color:var(--cio-border)!important;
}
[data-testid="stForm"] label,
[data-testid="stForm"] p,
[data-testid="stVerticalBlockBorderWrapper"] p,
[data-testid="stVerticalBlockBorderWrapper"] strong{color:var(--cio-text)!important}
input,textarea,[data-baseweb="select"]>div{
  background:#fff!important;
  color:#17202b!important;
  border-color:#bfc7d1!important;
}
input::placeholder,textarea::placeholder{color:#7b8794!important;opacity:1!important}
/* Metricas, tabelas e mensagens */
[data-testid="stMetric"]{background:#fff!important;border:1px solid var(--cio-border)!important;border-radius:12px!important;padding:12px!important}
[data-testid="stMetricLabel"],[data-testid="stMetricValue"],[data-testid="stMetricDelta"]{color:var(--cio-text)!important}
[data-testid="stDataFrame"],[data-testid="stTable"]{background:#fff!important;color:var(--cio-text)!important;border-radius:12px!important;overflow:hidden!important}
/* Segmented controls */
[data-testid="stSegmentedControl"] button,[data-testid="stSegmentedControl"] label{background:#fff!important;color:#202633!important;border-color:#cfd5dc!important}
[data-testid="stSegmentedControl"] button[aria-pressed="true"],[data-testid="stSegmentedControl"] label:has(input:checked){background:#a50f18!important;color:#fff!important;border-color:#d92832!important}
/* Botoes */
.stButton button,[data-testid="stFormSubmitButton"] button,[data-testid="stDownloadButton"] button{font-weight:750!important}
.stButton button[kind="primary"],[data-testid="stFormSubmitButton"] button[kind="primary"]{background:#c93449!important;color:#fff!important;border-color:#c93449!important}
/* Sidebar clara e legivel */
[data-testid="stSidebar"]{background:#fff!important;color:#202633!important;border-right:1px solid #dde2e8!important}
[data-testid="stSidebar"] *{color:#202633}
[data-testid="stSidebar"] [data-testid="stRadio"] label:has(input:checked){background:#fdebed!important;color:#b51820!important}
/* Mantem o dashboard executivo escuro e com alto contraste */
.cio-title,.cio-subtitle,.cio-update,.cio-kpi,.cio-panel-title,.cio-panel-sub,.cio-alert,.project-grid,.financial-row,.cio-highlight,.chart-panel-title{color:#fff!important}
.cio-title{color:#fff!important}.cio-subtitle{color:#d8cfd1!important}.cio-update{color:#d7ced0!important}.cio-update strong{color:#fff!important}
.cio-kpi-title,.cio-kpi-value{color:#fff!important}.cio-kpi-unit{color:#d2c7c9!important}.cio-panel-sub{color:#c4b8bb!important}
.cio-alert small{color:#c7bcc0!important}.project-grid b,.project-grid small{color:#d8cfd1!important}.financial-row small{color:#c9bec1!important}.cio-highlight{color:#f0ecee!important}
/* Titulos da Visao Geral sobre fundo escuro */
body:has(.cio-title) [data-testid="stAppViewContainer"]{background:radial-gradient(circle at 20% 0%,#39080d 0,#16090d 30%,#070708 72%)!important}
body:has(.cio-title) [data-testid="stMainBlockContainer"],body:has(.cio-title) .block-container{color:#fff!important}
body:has(.cio-title) [data-testid="stMainBlockContainer"] h1,
body:has(.cio-title) [data-testid="stMainBlockContainer"] h2,
body:has(.cio-title) [data-testid="stMainBlockContainer"] h3,
body:has(.cio-title) [data-testid="stMainBlockContainer"] p,
body:has(.cio-title) [data-testid="stMainBlockContainer"] label{color:#f3f5f7!important}
/* Select do periodo na visao geral */
body:has(.cio-title) [data-baseweb="select"]>div{background:#171015!important;color:#fff!important;border-color:#5b343a!important}
body:has(.cio-title) [data-baseweb="select"] span,body:has(.cio-title) [data-baseweb="select"] svg{color:#fff!important;fill:#fff!important}
/* Links sempre visiveis */
a{color:#0b63ce!important}body:has(.cio-title) a{color:#72b7ff!important}
/* Minha conta */
body:not(:has(.cio-title)) [data-testid="stVerticalBlockBorderWrapper"]{box-shadow:0 4px 14px rgba(31,41,55,.06)!important}
@media(max-width:580px){
  [data-testid="stAppViewContainer"] h1{font-size:1.55rem!important}
  [data-testid="stAppViewContainer"] h2{font-size:1.25rem!important}
}

/* Filtro de CD na Visao Geral */
body:has(.cio-title) [data-testid="stSelectbox"] label{color:#f3f5f7!important;font-size:.65rem!important;font-weight:700!important}
body:has(.cio-title) [data-testid="stSelectbox"]>div>div{min-height:40px!important}

/* Tabela executiva dos CDs, alinhada ao restante da Visao Geral */
.cio-table-wrap{background:#100c0f;border:1px solid #55262c;border-radius:14px;overflow:auto;max-height:300px;margin-top:8px}
.cio-table{width:100%;border-collapse:collapse;color:#f4f5f7;font-size:.76rem}
.cio-table th{position:sticky;top:0;background:#1c1115;color:#f7f7f8;text-align:left;padding:11px 10px;border-bottom:1px solid #643038;font-size:.67rem;text-transform:uppercase;letter-spacing:.35px}
.cio-table td{padding:11px 10px;border-bottom:1px solid #352126;color:#e9e4e5}
.cio-table tbody tr:hover{background:#211318}
.cio-table tbody tr:last-child td{border-bottom:0}

/* Destaques executivos de melhor e pior desempenho */
.decision-card{border:1px solid #593038;background:linear-gradient(145deg,#160d11,#0d0a0c);border-radius:13px;padding:12px 14px;min-height:88px;box-shadow:0 8px 22px rgba(0,0,0,.22)}
.decision-card span{display:block;font-size:.64rem;font-weight:900;letter-spacing:.45px;color:#cfc4c7;margin-bottom:5px}.decision-card strong{display:block;color:#fff;font-size:1.02rem;line-height:1.15}.decision-card small{display:block;color:#bfb3b7;font-size:.68rem;margin-top:5px}.decision-card.good{border-left:4px solid #1dc267;background:linear-gradient(145deg,rgba(24,184,75,.14),#0d0a0c 55%)}.decision-card.bad{border-left:4px solid #ef3340;background:linear-gradient(145deg,rgba(239,51,64,.16),#0d0a0c 55%)}
.cio-table-wrap.enhanced{max-height:330px}.cio-table tr.ok{background:rgba(29,194,103,.06)}.cio-table tr.risk{background:rgba(239,51,64,.08)}.cio-table tr.warn{background:rgba(245,166,35,.05)}
.status-pill{display:inline-block;border-radius:999px;padding:3px 8px;font-size:.58rem;font-weight:900;white-space:nowrap}.status-pill.ok{background:rgba(29,194,103,.18);color:#5ce694}.status-pill.warn{background:rgba(245,166,35,.18);color:#ffc75d}.status-pill.risk{background:rgba(239,51,64,.18);color:#ff7580}
@media(max-width:900px){.decision-card{min-height:78px;margin-bottom:5px}.cio-table{min-width:650px}}
</style>
<div class="portal-head portal-head-base">
  <div class="portal-brand">
    <img class="portal-logo-claro" src="https://raw.githubusercontent.com/TorredaLogistica/central-inteligencia-operacional-piloto/main/logo_claro.png?v=2" alt="Claro">
    <span class="portal-sep"></span>
    <div class="portal-title">Central de Inteligência Operacional</div>
    <span class="portal-sep"></span>
    <img class="portal-logo-logistica" src="https://raw.githubusercontent.com/TorredaLogistica/central-inteligencia-operacional-piloto/main/logo_logistica.png?v=2" alt="Logística">
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
    url_email_pendente = st.session_state.pop("outlook_email_pendente", None)
    if url_email_pendente:
        url_email_js = json.dumps(url_email_pendente)
        components_html(
            f"<script>window.parent.open({url_email_js}, '_blank', 'noopener,noreferrer');</script>",
            height=0,
        )
        st.info("O e-mail de retorno foi preparado no Outlook Web. Revise a mensagem e clique em Enviar.")
        st.link_button(
            "Abrir e-mail no Outlook Web",
            url_email_pendente,
            use_container_width=True,
        )
    st.markdown(f'<div class="pendencia">{len(pendentes)} solicitação(ões) pendente(s)</div>', unsafe_allow_html=True)
    st.caption("Datas e horários apresentados no fuso de Brasília (America/Sao_Paulo).")
    st.write("")

    aba_pendencias, aba_usuarios, aba_acessos, aba_importacao = st.tabs([
        f"Solicitações pendentes ({len(pendentes)})",
        f"Usuários cadastrados ({len(usuarios_admin)})",
        "Acessos aos indicadores",
        "Backup de usuários",
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
                    st.session_state.outlook_email_pendente = gerar_url_email_status_solicitacao(item, True, obs)
                    st.rerun()
                if c2.button("Rejeitar", key=f"rejeitar_{item['id']}", use_container_width=True):
                    decidir_solicitacao(item["id"], False, st.session_state.admin_logado, obs)
                    st.session_state.outlook_email_pendente = gerar_url_email_status_solicitacao(item, False, obs)
                    st.rerun()

    with aba_usuarios:
        nome_relacao_usuarios = f"Relacao_Usuarios_Torre_{datetime.now(FUSO_BRASILIA).strftime('%Y-%m-%d')}.xlsx"
        st.download_button(
            "Baixar relação de usuários (.xlsx)",
            data=gerar_relacao_usuarios_xlsx(usuarios_admin),
            file_name=nome_relacao_usuarios,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="baixar_relacao_usuarios_xlsx",
        )
        st.caption(
            "A exportação inclui todos os cadastros atuais com nome, usuário, e-mail, perfil, "
            "status e datas de criação, aprovação e atualização."
        )
        st.divider()
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

    with aba_acessos:
        exibir_relatorio_acessos(usuarios_admin)

    with aba_importacao:
        st.subheader("Backup de usuários")
        nome_backup = f"backup_usuarios_completo_{datetime.now(FUSO_BRASILIA).strftime('%Y%m%d_%H%M%S')}.json"
        st.download_button(
            "Exportar backup completo de usuários",
            exportar_backup_completo_usuarios(),
            nome_backup,
            "application/json",
            use_container_width=True,
        )
        st.divider()
        st.info(
            "Selecione o backup JSON completo. A restauração inclui usuários novos e atualiza os já existentes, "
            "preservando nomes, credenciais, status, perfil e datas administrativas."
        )
        arquivo_json = st.file_uploader(
            "Selecione o backup_usuarios_completo.json",
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
                    f"Restauração concluída. Analisados: {resultado['analisados']} | "
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


    if st.button("Sair da administração"):
        st.session_state.admin_logado = None
        st.rerun()
    st.stop()

if st.session_state.usuario_logado:
    exibir_portal_usuario()
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
            st.session_state.pagina_portal_usuario = "🏠 Visão Geral"
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
            criar_solicitacao_cadastro(nome, usuario, email, senha)
            st.success(
                "Solicitação registrada. O acesso será liberado somente após aprovação do administrador, "
                "e o retorno ocorrerá por e-mail em até 24 horas úteis."
            )
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
                st.success(
                    "Se os dados corresponderem a um cadastro ativo, a solicitação será encaminhada para aprovação, "
                    "e o retorno ocorrerá por e-mail em até 24 horas úteis."
                )
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
