import streamlit as st
import pandas as pd
import datetime
import calendar
import smtplib
import ssl
import openpyxl
import json
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication

# -----------------------------------------------------------------------------
# CONFIGURAÇÕES DE ARCHIVO E EMAIL
# -----------------------------------------------------------------------------
SMTP_SERVER = "mail.feital.com.br"
SMTP_PORT = 587
EMAIL_REMETENTE = "carlos.rene@feital.com.br"
SENHA_REMETENTE = "S@turno2026"
EMAIL_SETOR_COMPRAS = "carlos.rene@feital.com.br"

EXCEL_PATH = 'CONTRATOS_2026.xlsx'
CONFIG_PATH = 'config_colunas.json'

# Mapeamento de meses
MESES_LISTA = [
    'JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO',
    'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'
]

# Larguras padrão iniciais (caso não exista JSON salvo)
LARGURAS_PADRAO = {
    "LOCAL": 110,
    "FATURAMENTO": 170,
    "PRESTADOR": 180,
    "CONTA": 120,
    "SERVIÇO": 320,
    "DT. EMISSÃO NF": 160,
    "VENC.": 110,
    "VALOR_MES": 160,
    "STATUS": 120
}

# -----------------------------------------------------------------------------
# GERENCIAMENTO DE CONFIGURAÇÃO DE LARGURAS (PERSISTÊNCIA)
# -----------------------------------------------------------------------------
def carregar_larguras():
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return LARGURAS_PADRAO.copy()

def salvar_larguras(larguras_dict):
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(larguras_dict, f, indent=4)
        return True
    except Exception as e:
        st.error(f"Erro ao salvar configuração de colunas: {e}")
        return False

# -----------------------------------------------------------------------------
# FUNÇÕES DE FORMATAÇÃO E TRATAMENTO DE DADOS
# -----------------------------------------------------------------------------
def formata_br(valor):
    try:
        val_float = float(valor)
        return f"R$ {val_float:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return "R$ 0,00"

def converter_para_numero(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace('R$', '').replace(' ', '')
    if ',' in s:
        s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0

def calcular_data_valida(dia_param):
    hoje = datetime.date.today()
    try:
        if pd.isna(dia_param) or dia_param is None:
            return hoje
        dia = int(float(dia_param))
        if 1 <= dia <= 31:
            max_dias = calendar.monthrange(hoje.year, hoje.month)[1]
            dia_ajustado = min(dia, max_dias)
            return datetime.date(hoje.year, hoje.month, dia_ajustado)
    except (ValueError, TypeError):
        pass
    return hoje

def enviar_email(destino, assunto, corpo_html, anexo=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_REMETENTE
        msg['To'] = destino
        msg['Subject'] = assunto
        msg.attach(MIMEText(corpo_html, 'html'))

        if anexo is not None:
            part = MIMEApplication(anexo.getvalue(), Name=anexo.name)
            part['Content-Disposition'] = f'attachment; filename="{anexo.name}"'
            msg.attach(part)

        context = ssl.create_default_context()
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.ehlo()
            server.starttls(context=context)
            server.ehlo()
            server.login(EMAIL_REMETENTE, SENHA_REMETENTE)
            server.send_message(msg)
        return True
    except Exception as e:
        st.error(f"Erro ao enviar e-mail: {e}")
        return False

# -----------------------------------------------------------------------------
# CONFIGURAÇÃO DA PÁGINA
# -----------------------------------------------------------------------------
st.set_page_config(page_title="Gestão de Pagamentos & Contratos", layout="wide")

st.markdown("""
    <style>
    [data-testid="stTable"] th, [data-testid="stDataEditor"] th {
        white-space: nowrap !important;
        text-overflow: clip !important;
        font-weight: bold !important;
    }
    </style>
""", unsafe_allow_html=True)

def carregar_dados():
    df_consolidado = pd.read_excel(EXCEL_PATH, sheet_name='CONSOLIDADO')
    df_licencas = pd.read_excel(EXCEL_PATH, sheet_name='LICENÇAS_ANUAL')
    
    for col in MESES_LISTA:
        if col in df_consolidado.columns:
            df_consolidado[col] = df_consolidado[col].apply(converter_para_numero)
            
    return df_consolidado, df_licencas

try:
    if "df_consolidado" not in st.session_state:
        st.session_state.df_consolidado, st.session_state.df_licencas = carregar_dados()
    
    df_consolidado = st.session_state.df_consolidado
    df_licencas = st.session_state.df_licencas
except Exception as e:
    st.error(f"Erro ao carregar a planilha {EXCEL_PATH}: {e}")
    st.stop()

# Carrega a configuração salva de larguras
if "larguras" not in st.session_state:
    st.session_state.larguras = carregar_larguras()

# -----------------------------------------------------------------------------
# MENU DE NAVEGAÇÃO LATERAL
# -----------------------------------------------------------------------------
st.sidebar.title("📌 Navegação")
pagina = st.sidebar.radio(
    "Selecione o Módulo:",
    ["📋 Faturas Mensais", "📜 Licenças Anuais"],
    key="nav_menu"
)

# =============================================================================
# PÁGINA 1: FATURAS MENSAIS (FORNECEDORES)
# =============================================================================
if pagina == "📋 Faturas Mensais":
    st.title("📊 Gestão e Controle de Pagamentos de Fornecedores")
    st.subheader("Faturas Mensais por Prestador")
    
    # 1. FILTROS SUPERIORES (MÊS DINÂMICO BASEADO NO MÊS ATUAL)
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        prestadores_validos = sorted([p for p in df_consolidado['PRESTADOR'].dropna().unique() if str(p).strip().upper() != 'TOTAL'])
        prestadores = ["Todos"] + prestadores_validos
        prestador_sel = st.selectbox("Filtrar por Prestador:", prestadores, key="sb_prestador")
    with col_f2:
        mes_atual_idx = datetime.date.today().month - 1  # Detecta o mês atual automaticamente
        mes_sel = st.selectbox("Selecione o Mês de Referência:", MESES_LISTA, index=mes_atual_idx, key="sb_mes")

    # Painel Expansível para Ajustar e Salvar o Tamanho das Colunas
    with st.expander("⚙️ Personalizar Largura das Colunas"):
        st.markdown("Ajuste as larguras abaixo e clique em **Salvar Layout**. Na próxima vez que abrir o programa, as colunas estarão exatamente com esse tamanho!")
        
        col_c1, col_c2, col_c3, col_c4 = st.columns(4)
        with col_c1:
            st.session_state.larguras["FATURAMENTO"] = st.number_input("FATURAMENTO (px)", value=int(st.session_state.larguras.get("FATURAMENTO", 170)), step=10)
            st.session_state.larguras["PRESTADOR"] = st.number_input("PRESTADOR (px)", value=int(st.session_state.larguras.get("PRESTADOR", 180)), step=10)
        with col_c2:
            st.session_state.larguras["SERVIÇO"] = st.number_input("SERVIÇO (px)", value=int(st.session_state.larguras.get("SERVIÇO", 320)), step=10)
            st.session_state.larguras["DT. EMISSÃO NF"] = st.number_input("DT. EMISSÃO NF (px)", value=int(st.session_state.larguras.get("DT. EMISSÃO NF", 160)), step=10)
        with col_c3:
            st.session_state.larguras["VENC."] = st.number_input("VENC. (px)", value=int(st.session_state.larguras.get("VENC.", 110)), step=10)
            st.session_state.larguras["VALOR_MES"] = st.number_input("VALOR DO MÊS (px)", value=int(st.session_state.larguras.get("VALOR_MES", 160)), step=10)
        with col_c4:
            st.session_state.larguras["LOCAL"] = st.number_input("LOCAL (px)", value=int(st.session_state.larguras.get("LOCAL", 110)), step=10)
            st.session_state.larguras["CONTA"] = st.number_input("CONTA (px)", value=int(st.session_state.larguras.get("CONTA", 120)), step=10)

        if st.button("💾 Salvar Layout das Colunas no Computador", key="btn_salvar_layout"):
            if salvar_larguras(st.session_state.larguras):
                st.success("✅ Layout de colunas salvo com sucesso! Ele será mantido sempre que você abrir o programa.")

    colunas_fixas = ['LOCAL', 'FATURAMENTO', 'PRESTADOR', 'CONTA', 'SERVIÇO', 'DT. EMISSÃO NF', 'VENC.']

    idx_total = df_consolidado[df_consolidado['LOCAL'].astype(str).str.upper() == 'TOTAL'].index
    if not idx_total.empty:
        pos_total = idx_total[0]
        df_fornecedores = df_consolidado.iloc[:pos_total].copy()
    else:
        df_fornecedores = df_consolidado.copy()

    if prestador_sel != "Todos":
        df_fornecedores = df_fornecedores[df_fornecedores['PRESTADOR'] == prestador_sel]

    df_fornecedores['STATUS'] = df_fornecedores[mes_sel].apply(
        lambda v: '✅ OK' if converter_para_numero(v) > 0 else '⏳ Pendente'
    )

    colunas_exibir = [c for c in colunas_fixas if c in df_fornecedores.columns] + [mes_sel, 'STATUS']

    # 3. TABELA EDITÁVEL USANDO AS LARGURAS SALVAS NO ARQUIVO LOCAL
    l = st.session_state.larguras
    df_editado_mes = st.data_editor(
        df_fornecedores[colunas_exibir],
        use_container_width=True,
        hide_index=True,
        column_config={
            "LOCAL": st.column_config.TextColumn("LOCAL", width=l.get("LOCAL", 110)),
            "FATURAMENTO": st.column_config.TextColumn("FATURAMENTO", width=l.get("FATURAMENTO", 170)),
            "PRESTADOR": st.column_config.TextColumn("PRESTADOR", width=l.get("PRESTADOR", 180)),
            "CONTA": st.column_config.TextColumn("CONTA", width=l.get("CONTA", 120)),
            "SERVIÇO": st.column_config.TextColumn("SERVIÇO", width=l.get("SERVIÇO", 320)),
            "DT. EMISSÃO NF": st.column_config.NumberColumn("DT. EMISSÃO NF", step=1, width=l.get("DT. EMISSÃO NF", 160)),
            "VENC.": st.column_config.NumberColumn("Dia Venc.", step=1, width=l.get("VENC.", 110)),
            mes_sel: st.column_config.NumberColumn(
                f"Valor ({mes_sel})",
                format="R$ %.2f",
                step=0.01,
                width=l.get("VALOR_MES", 160)
            ),
            "STATUS": st.column_config.TextColumn("STATUS", disabled=True, width=l.get("STATUS", 120))
        },
        key=f"editor_mensal_{mes_sel}_{prestador_sel}"
    )

    # Captura edições de valor na memória
    houve_alteracao = False
    for idx, row in df_editado_mes.iterrows():
        val_digitado = converter_para_numero(row[mes_sel])
        val_antigo = converter_para_numero(st.session_state.df_consolidado.loc[idx, mes_sel])
        if val_digitado != val_antigo:
            st.session_state.df_consolidado.loc[idx, mes_sel] = val_digitado
            houve_alteracao = True

    if houve_alteracao:
        st.rerun()

    # 4. TOTAIS E LINHA DE TOTALIZAÇÃO
    valores_limpos = df_editado_mes[mes_sel].apply(converter_para_numero)
    total_mes = valores_limpos.sum()
    qtd_faturas = (valores_limpos > 0).sum()
    media_faturas = total_mes / qtd_faturas if qtd_faturas > 0 else 0.0

    df_linha_total = pd.DataFrame([{
        'LOCAL': 'TOTAL',
        'FATURAMENTO': '',
        'PRESTADOR': '',
        'CONTA': '',
        'SERVIÇO': 'SOMA DOS FORNECEDORES DO MÊS',
        'DT. EMISSÃO NF': '',
        'VENC.': '',
        mes_sel: total_mes,
        'STATUS': '📊 TOTAL'
    }])

    st.dataframe(
        df_linha_total[colunas_exibir],
        use_container_width=True,
        hide_index=True,
        column_config={
            "LOCAL": st.column_config.TextColumn("LOCAL", width=l.get("LOCAL", 110)),
            "FATURAMENTO": st.column_config.TextColumn("FATURAMENTO", width=l.get("FATURAMENTO", 170)),
            "PRESTADOR": st.column_config.TextColumn("PRESTADOR", width=l.get("PRESTADOR", 180)),
            "CONTA": st.column_config.TextColumn("CONTA", width=l.get("CONTA", 120)),
            "SERVIÇO": st.column_config.TextColumn("SERVIÇO", width=l.get("SERVIÇO", 320)),
            "DT. EMISSÃO NF": st.column_config.TextColumn("DT. EMISSÃO NF", width=l.get("DT. EMISSÃO NF", 160)),
            "VENC.": st.column_config.TextColumn("Dia Venc.", width=l.get("VENC.", 110)),
            mes_sel: st.column_config.NumberColumn(
                f"Valor Total ({mes_sel})",
                format="R$ %.2f",
                width=l.get("VALOR_MES", 160)
            ),
            "STATUS": st.column_config.TextColumn("STATUS", width=l.get("STATUS", 120))
        }
    )

    # Cards Resumo
    st.markdown("### 🧮 Resumo de Valores e Soma do Mês")
    col_s1, col_s2, col_s3 = st.columns(3)
    col_s1.metric(f"💰 SOMA TOTAL ({mes_sel})", formata_br(total_mes))
    col_s2.metric("📄 Faturas com Valor", f"{qtd_faturas} lançamentos")
    col_s3.metric("📊 Média por Fatura", formata_br(media_faturas))

    st.markdown("---")

    # Botão de Salvamento
    if st.button("💾 Salvar Alterações na Planilha Excel", key="btn_salvar_mensal"):
        try:
            wb = openpyxl.load_workbook(EXCEL_PATH)
            ws = wb['CONSOLIDADO']

            header = [cell.value for cell in ws[1]]
            if mes_sel in header:
                col_idx = header.index(mes_sel) + 1

                for idx, row in df_editado_mes.iterrows():
                    val_num = converter_para_numero(row[mes_sel])
                    excel_row = idx + 2
                    ws.cell(row=excel_row, column=col_idx, value=val_num)

                for r in range(2, ws.max_row + 1):
                    local_val = str(ws.cell(row=r, column=1).value).strip().upper()
                    if local_val == 'TOTAL':
                        ws.cell(row=r, column=col_idx, value=total_mes)
                        break

                wb.save(EXCEL_PATH)
                st.success(f"✅ Valores do mês de {mes_sel} e TOTAL ({formata_br(total_mes)}) salvos com sucesso!")
            else:
                st.error(f"Coluna {mes_sel} não encontrada na planilha Excel.")
        except Exception as e:
            st.error(f"Erro ao salvar na planilha Excel: {e}")

    st.markdown("---")
    
    # =========================================================================
    # FORMULÁRIO DINÂMICO DE SOLICITAÇÃO (E-MAIL COM NOTA FISCAL ANEXA)
    # =========================================================================
    st.subheader("📩 Gerar Solicitação de Pedido de Compra (Setor de Compras)")

    def preencher_campos_fornecedor():
        p_selecionado = st.session_state.get("form_prestador_sel")
        if not p_selecionado:
            return

        linha_forn = df_consolidado[df_consolidado['PRESTADOR'] == p_selecionado]
        
        if not linha_forn.empty:
            dados = linha_forn.iloc[0]
            st.session_state.form_desc = str(dados.get('SERVIÇO', '')).replace('nan', '')
            
            mes_ref = st.session_state.get('sb_mes', 'AGOSTO')
            val_mes = converter_para_numero(dados.get(mes_ref, 0))
            if val_mes == 0:
                for m in reversed(MESES_LISTA):
                    if m in dados and converter_para_numero(dados[m]) > 0:
                        val_mes = converter_para_numero(dados[m])
                        break
            st.session_state.form_valor = float(val_mes)
            
            st.session_state.form_dt_emissao = calcular_data_valida(dados.get('DT. EMISSÃO NF', None))
            st.session_state.form_dt_venc = calcular_data_valida(dados.get('VENC.', None))

    if "form_prestador_sel" not in st.session_state and prestadores_validos:
        st.session_state.form_prestador_sel = prestadores_validos[0]
        preencher_campos_fornecedor()

    with st.container(border=True):
        st.selectbox(
            "Prestador / Fornecedor",
            prestadores_validos,
            key="form_prestador_sel",
            on_change=preencher_campos_fornecedor
        )

        col_p1, col_p2, col_p3 = st.columns([2, 1.5, 1])
        with col_p1:
            st.text_input("Descrição do Serviço / Item", key="form_desc")
        with col_p2:
            st.number_input("Valor da Fatura (R$)", step=10.0, format="%.2f", key="form_valor")
        with col_p3:
            st.text_input("Número da NF / Fatura", value="NF-157896", key="form_nf")

        col_p4, col_p5 = st.columns(2)
        with col_p4:
            st.date_input("Data de Emissão da NF", format="DD/MM/YYYY", key="form_dt_emissao")
        with col_p5:
            st.date_input("Data de Vencimento", format="DD/MM/YYYY", key="form_dt_venc")

        st.text_area("Observações Adicionais para Compras", value="segue nota fiscal em anexo", height=100, key="form_obs")

        arquivo_nf = st.file_uploader(
            "📎 Anexar Nota Fiscal / Comprovante (PDF, PNG, JPG, XML, ZIP)",
            type=["pdf", "png", "jpg", "jpeg", "xml", "zip"],
            key="form_anexo"
        )

        if st.button("✉️ Enviar para Setor de Compras", key="btn_enviar_solicitacao"):
            if arquivo_nf is None:
                st.warning("⚠️ **Atenção:** O Setor de Compras exige o envio da Nota Fiscal! Por favor, anexe o arquivo acima antes de enviar.")
            else:
                f_nome = st.session_state.form_prestador_sel
                f_desc = st.session_state.form_desc
                f_val = st.session_state.form_valor
                f_nf = st.session_state.form_nf
                f_dt_em = st.session_state.form_dt_emissao.strftime('%d/%m/%Y')
                f_dt_vc = st.session_state.form_dt_venc.strftime('%d/%m/%Y')
                f_obs = st.session_state.form_obs

                corpo_email = f"""
                <h3>Solicitação de Pedido de Compra</h3>
                <p>Prezados,</p>
                <p>Por gentileza, gerar o Pedido de Compra referente aos dados abaixo:</p>
                <ul>
                    <li><b>Fornecedor / Prestador:</b> {f_nome}</li>
                    <li><b>Serviço:</b> {f_desc}</li>
                    <li><b>Número da NF:</b> {f_nf}</li>
                    <li><b>Data de Emissão:</b> {f_dt_em}</li>
                    <li><b>Data de Vencimento:</b> {f_dt_vc}</li>
                    <li><b>Valor:</b> {formata_br(f_val)}</li>
                    <li><b>Observações:</b> {f_obs}</li>
                    <li><b>Anexo:</b> Nota Fiscal vinculada na mensagem.</li>
                </ul>
                <br>
                <p><i>E-mail enviado automaticamente pelo Sistema de Gestão de Contratos.</i></p>
                """
                
                sucesso = enviar_email(
                    destino=EMAIL_SETOR_COMPRAS,
                    assunto=f"Pedido de Compra | Nota Fiscal - {f_nome}",
                    corpo_html=corpo_email,
                    anexo=arquivo_nf
                )
                if sucesso:
                    st.success(f"✅ Solicitação e Nota Fiscal (**{arquivo_nf.name}**) enviadas com sucesso para **{f_nome}** no valor de **{formata_br(f_val)}**!")

# =============================================================================
# PÁGINA 2: LICENÇAS ANUAIS
# =============================================================================
elif pagina == "📜 Licenças Anuais":
    st.title("📜 Gestão de Licenças Anuais e Renovações")
    st.markdown("Visualização, edição e atualização dos valores das licenças anuais contratadas.")
    st.markdown("---")

    for col in ['VALOR PAGO ANO ANTERIOR', 'VALOR ATUAL']:
        if col in st.session_state.df_licencas.columns:
            st.session_state.df_licencas[col] = st.session_state.df_licencas[col].apply(converter_para_numero).astype(float)

    df_licencas_base = st.session_state.df_licencas.copy()
    mask_total = df_licencas_base['VENCIMENTO'].astype(str).str.upper() == 'TOTAL'
    df_lic_itens = df_licencas_base[~mask_total].copy() if mask_total.any() else df_licencas_base.copy()

    l = st.session_state.larguras
    df_lic_editado = st.data_editor(
        df_lic_itens,
        use_container_width=True,
        hide_index=True,
        column_config={
            "LOCAL": st.column_config.TextColumn("LOCAL", width=l.get("LOCAL", 110)),
            "FATURAMENTO": st.column_config.TextColumn("FATURAMENTO", width=l.get("FATURAMENTO", 170)),
            "PRESTADOR": st.column_config.TextColumn("PRESTADOR", width=l.get("PRESTADOR", 180)),
            "SERVIÇO": st.column_config.TextColumn("SERVIÇO", width=l.get("SERVIÇO", 320)),
            "VALOR PAGO ANO ANTERIOR": st.column_config.NumberColumn("VALOR PAGO ANO ANTERIOR", format="R$ %.2f", step=0.01, width=200),
            "VALOR ATUAL": st.column_config.NumberColumn("VALOR ATUAL", format="R$ %.2f", step=0.01, width=160),
            "VENC.": st.column_config.NumberColumn("Dia Venc.", step=1, width=l.get("VENC.", 110)),
            "DT. EMISSÃO NF": st.column_config.NumberColumn("Dia Emissão", step=1, width=130)
        },
        key="editor_licencas_pagina_separada"
    )

    houve_alteracao_lic = False
    for idx, row in df_lic_editado.iterrows():
        val_ant_dig = converter_para_numero(row['VALOR PAGO ANO ANTERIOR'])
        val_atu_dig = converter_para_numero(row['VALOR ATUAL'])
        
        val_ant_ori = converter_para_numero(st.session_state.df_licencas.loc[idx, 'VALOR PAGO ANO ANTERIOR'])
        val_atu_ori = converter_para_numero(st.session_state.df_licencas.loc[idx, 'VALOR ATUAL'])
        
        if (val_ant_dig != val_ant_ori) or (val_atu_dig != val_atu_ori):
            st.session_state.df_licencas.loc[idx, 'VALOR PAGO ANO ANTERIOR'] = val_ant_dig
            st.session_state.df_licencas.loc[idx, 'VALOR ATUAL'] = val_atu_dig
            houve_alteracao_lic = True

    if houve_alteracao_lic:
        st.rerun()

    tot_anterior = df_lic_editado['VALOR PAGO ANO ANTERIOR'].apply(converter_para_numero).sum()
    tot_atual = df_lic_editado['VALOR ATUAL'].apply(converter_para_numero).sum()

    df_linha_total_licencas = pd.DataFrame([{
        'LOCAL': '',
        'FATURAMENTO': '',
        'PRESTADOR': '',
        'TIPO DE ENVIO': '',
        'SERVIÇO': 'TOTAL DAS LICENÇAS',
        'DT. EMISSÃO NF': '',
        'VENC.': '',
        'VENCIMENTO': 'TOTAL',
        'VALOR PAGO ANO ANTERIOR': tot_anterior,
        'VALOR ATUAL': tot_atual
    }])

    st.dataframe(
        df_linha_total_licencas,
        use_container_width=True,
        hide_index=True,
        column_config={
            "VALOR PAGO ANO ANTERIOR": st.column_config.NumberColumn("Total Pago Anterior", format="R$ %.2f", width=200),
            "VALOR ATUAL": st.column_config.NumberColumn("Total Atual", format="R$ %.2f", width=160)
        }
    )

    st.markdown("### 📊 Resumo dos Totais")
    c1, c2 = st.columns(2)
    c1.metric("💰 Total Pago Ano Anterior", formata_br(tot_anterior))
    c2.metric("💳 Total Atual", formata_br(tot_atual))

    st.markdown("<br>", unsafe_allow_html=True)

    if st.button("💾 Salvar Alterações nas Licenças Anuais", key="btn_salvar_licencas_pagina"):
        try:
            wb = openpyxl.load_workbook(EXCEL_PATH)
            ws = wb['LICENÇAS_ANUAL']

            for r_idx, row in df_lic_editado.iterrows():
                for c_idx, val in enumerate(row):
                    ws.cell(row=r_idx + 2, column=c_idx + 1, value=val)

            for r in range(2, ws.max_row + 1):
                cell_venc = str(ws.cell(row=r, column=8).value).strip().upper()
                if cell_venc == 'TOTAL':
                    ws.cell(row=r, column=9, value=tot_anterior)
                    ws.cell(row=r, column=10, value=tot_atual)
                    break

            wb.save(EXCEL_PATH)
            st.success("✅ Alterações salvas com sucesso na aba LICENÇAS_ANUAL!")
        except Exception as e:
            st.error(f"Erro ao salvar as licenças na planilha Excel: {e}")