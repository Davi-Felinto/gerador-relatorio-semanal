"""
app.py
------
Interface Streamlit para gerar o relatorio semanal (Canal Educacao)
sem precisar rodar nada por linha de comando.

Rodar com:
    streamlit run app.py
"""

import io

import streamlit as st

from report_core import generate_report_html, list_data_sheets

st.set_page_config(page_title="Gerador de Relatório Semanal", page_icon="📧", layout="centered")

st.title("📧 Gerador de Relatório Semanal")
st.caption("Suba a planilha da semana, escolha a aba e gere o HTML do email.")

uploaded_file = st.file_uploader("Planilha (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # Guarda os bytes originais para poder reabrir o arquivo varias vezes
    # (openpyxl consome o stream a cada leitura, e listamos abas + geramos
    # o relatorio em passos separados).
    file_bytes = uploaded_file.getvalue()

    try:
        sheet_names = list_data_sheets(io.BytesIO(file_bytes))
    except Exception as e:
        st.error(f"Não consegui ler as abas do arquivo: {e}")
        sheet_names = []

    if not sheet_names:
        st.warning(
            "Não encontrei nenhuma aba de dados nesse arquivo "
            "(além de 'Config'/'KPIs', que são reservadas)."
        )
    else:
        sheet_name = st.selectbox(
            "Aba da semana",
            options=sheet_names,
            index=len(sheet_names) - 1,  # a mais recente costuma ser a ultima aba
            help="A aba com o período da semana, ex: '31-08 a 04-09'.",
        )

        if st.button("Gerar relatório", type="primary"):
            try:
                html, resumo = generate_report_html(io.BytesIO(file_bytes), sheet_name)
            except Exception as e:
                st.error(f"Não consegui gerar o relatório: {e}")
            else:
                st.success(
                    f"Relatório gerado: {resumo['num_secoes']} seções · "
                    f"{resumo['total_concluido']} concluídas · "
                    f"{resumo['total_agendado']} agendadas · "
                    f"período {resumo['period_start']}–{resumo['period_end']}"
                )
                if not resumo["config_encontrada"]:
                    st.info(
                        "Não encontrei a aba 'Config' nesta planilha — período e "
                        "destaques foram preenchidos automaticamente. Revise antes "
                        "de enviar."
                    )
                if resumo["num_kpis"] == 0:
                    st.info("Aba 'KPIs' não encontrada ou vazia — a seção de indicadores ficará com aviso padrão.")

                st.subheader("Prévia")
                st.components.v1.html(html, height=600, scrolling=True)

                st.download_button(
                    label="⬇️ Baixar HTML",
                    data=html.encode("utf-8"),
                    file_name=f"relatorio_{sheet_name.replace(' ', '_')}.html",
                    mime="text/html",
                )
                st.caption(
                    "Depois de baixar: abra o HTML no navegador, selecione tudo "
                    "(Ctrl+A) e cole no corpo do email — não cole o código-fonte."
                )
else:
    st.info("Envie a planilha .xlsx para começar.")
