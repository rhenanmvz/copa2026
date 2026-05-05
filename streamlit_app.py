from __future__ import annotations

import io
import tempfile
from datetime import datetime
from pathlib import Path
from uuid import uuid4

import pandas as pd
import streamlit as st

from simulacao_copa_2026_completa import (
    DEFAULT_SIMULATION_PRESET,
    SIMULATION_PRESETS,
    get_simulation_preset,
    run_world_cup_simulation,
)


APP_TITLE = "Simulador da Copa do Mundo 2026"
MAX_CUPS_PER_USER = 10
LOCAL_WORKBOOK_PATH = Path(__file__).with_name("copachaveamento.xlsx")

MODE_DESCRIPTIONS = {
    "conservador": "Mais fiel ao favoritismo. Menos zebra e placares mais comportados.",
    "equilibrado": "O ponto de equilibrio entre coerencia estatistica e surpresa.",
    "caotico": "Mais oscilacao. Zebras aparecem com muito mais frequencia.",
}


def set_page_config() -> None:
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="🏆",
        layout="wide",
        initial_sidebar_state="expanded",
    )


def render_intro() -> None:
    st.title(APP_TITLE)
    st.caption(
        "Escolha um modo, rode ate 10 Copas por vez e veja como o ranking oficial muda a cada execucao."
    )

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Selecoes", "48")
    col2.metric("Grupos", "12")
    col3.metric("Jogos por Copa", "104")
    col4.metric("Limite por pessoa", f"{MAX_CUPS_PER_USER} Copas")


def render_sidebar() -> tuple[str, int, int | None, bytes | None]:
    st.sidebar.header("Configuracao da simulacao")
    if LOCAL_WORKBOOK_PATH.exists():
        st.sidebar.success(
            "Planilha local encontrada no projeto. O publico pode usar o app sem fazer upload."
        )
    else:
        st.sidebar.write(
            "Para publicar no Streamlit Cloud, o app aceita upload do arquivo Excel diretamente na pagina."
        )

    uploaded_file = st.sidebar.file_uploader(
        "Planilha base (.xlsx)",
        type=["xlsx"],
        help="Envie a planilha com as abas Selecoes, Saldo, Probabilidades e Chaveamento.",
    )

    mode = st.sidebar.radio(
        "Modo",
        options=list(SIMULATION_PRESETS.keys()),
        index=list(SIMULATION_PRESETS.keys()).index(DEFAULT_SIMULATION_PRESET),
        format_func=lambda key: f"{key.title()}",
        help="Cada modo muda o quanto a simulacao respeita favoritismo e abre espaco para zebra.",
    )

    cups = st.sidebar.slider(
        "Quantas Copas sortear nesta execucao?",
        min_value=1,
        max_value=MAX_CUPS_PER_USER,
        value=3,
        step=1,
        help="Limitada em 10 deixa o app rapido e evita sobrecarga.",
    )

    seed_text = st.sidebar.text_input(
        "Seed opcional",
        value="",
        help="Deixe em branco para gerar uma execucao nova e probabilistica.",
    )

    try:
        parsed_seed = int(seed_text.strip()) if seed_text.strip() else None
    except ValueError:
        st.sidebar.warning("A seed precisa ser um numero inteiro ou ficar em branco.")
        parsed_seed = None

    st.sidebar.markdown("---")
    st.sidebar.subheader("Resumo do modo")
    params = get_simulation_preset(mode)
    st.sidebar.write(MODE_DESCRIPTIONS.get(mode, ""))
    st.sidebar.dataframe(pd.DataFrame([params]), use_container_width=True, hide_index=True)

    if uploaded_file is not None:
        file_bytes = uploaded_file.getvalue()
    elif LOCAL_WORKBOOK_PATH.exists():
        file_bytes = LOCAL_WORKBOOK_PATH.read_bytes()
    else:
        file_bytes = None
    return mode, cups, parsed_seed, file_bytes


def save_uploaded_workbook(file_bytes: bytes) -> Path:
    temp_dir = Path(tempfile.mkdtemp(prefix="copa2026_upload_"))
    workbook_path = temp_dir / "copachaveamento.xlsx"
    workbook_path.write_bytes(file_bytes)
    return workbook_path


def run_simulation_from_upload(
    *,
    file_bytes: bytes,
    mode: str,
    cups: int,
    random_seed: int | None,
) -> dict[str, object]:
    input_path = save_uploaded_workbook(file_bytes)
    output_path = input_path.with_name(f"resultado_{mode}_{uuid4().hex[:8]}.xlsx")
    params = get_simulation_preset(mode)

    ranking_df, last_matches_df, last_groups_df, last_final_ranking_df, exported_path = run_world_cup_simulation(
        excel_path=input_path,
        output_path=output_path,
        n_simulations=cups,
        random_seed=random_seed,
        preset_name=mode,
        export_to_source_workbook=False,
        probability_concentration=params["probability_concentration"],
        probability_tilt=params["probability_tilt"],
        attack_weight=params["attack_weight"],
        defense_weight=params["defense_weight"],
        min_goals_lambda=params["min_goals_lambda"],
        max_goals_lambda=params["max_goals_lambda"],
    )

    return {
        "ranking_df": ranking_df,
        "last_matches_df": last_matches_df,
        "last_groups_df": last_groups_df,
        "last_final_ranking_df": last_final_ranking_df,
        "exported_path": exported_path,
        "output_bytes": exported_path.read_bytes(),
    }


def build_summary_cards(ranking_df: pd.DataFrame, final_ranking_df: pd.DataFrame, cups: int) -> None:
    top_champion = ranking_df.sort_values("Pct_Campeao", ascending=False).iloc[0]
    top_average = ranking_df.sort_values("Posicao_Media", ascending=True).iloc[0]
    last_champion = final_ranking_df.sort_values("posicao").iloc[0]

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Copas sorteadas", f"{cups}")
    col2.metric("Favorito ao titulo", str(top_champion["selecao"]), f"{float(top_champion['Pct_Campeao']):.2f}%")
    col3.metric("Melhor media final", str(top_average["selecao"]), f"{float(top_average['Posicao_Media']):.2f}")
    col4.metric("Campeao da ultima Copa", str(last_champion["selecao"]))


def render_results(result: dict[str, object], cups: int) -> None:
    ranking_df = result["ranking_df"]
    last_matches_df = result["last_matches_df"]
    last_groups_df = result["last_groups_df"]
    last_final_ranking_df = result["last_final_ranking_df"]
    output_bytes = result["output_bytes"]
    exported_path = result["exported_path"]

    st.success(
        f"Simulacao concluida. Modo {ranking_df.attrs.get('preset_name')}, seed {ranking_df.attrs.get('execution_seed')}."
    )
    st.caption(f"Horario da execucao: {ranking_df.attrs.get('execution_timestamp')}")

    build_summary_cards(ranking_df, last_final_ranking_df, cups)

    download_name = f"copa2026_streamlit_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    st.download_button(
        "Baixar Excel desta execucao",
        data=output_bytes,
        file_name=download_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.caption(f"Arquivo gerado internamente: {exported_path.name}")

    ranking_cols = [
        "Ranking_Oficial",
        "selecao",
        "grupo",
        "Posicao_Media",
        "Posicao_Mais_Provavel",
        "Pct_Campeao",
        "Pct_Finalista",
        "Pct_Semifinal",
        "Pct_Quartas",
        "Pct_Oitavas",
        "Pct_Mata_Mata",
    ]
    group_cols = ["grupo", "classificacao", "selecao", "P", "SG", "GP", "GC", "qualificado"]
    match_cols = ["match_no", "fase", "time_a", "time_b", "gols_a", "gols_b", "vencedor", "prob_base_a", "prob_sorteada_a"]

    tab1, tab2, tab3, tab4 = st.tabs(
        ["Ranking oficial", "Ultima Copa", "Grupos", "Jogos da ultima Copa"]
    )

    with tab1:
        st.dataframe(ranking_df[ranking_cols], use_container_width=True, hide_index=True)

    with tab2:
        st.dataframe(last_final_ranking_df, use_container_width=True, hide_index=True)

    with tab3:
        st.dataframe(last_groups_df[group_cols], use_container_width=True, hide_index=True)

    with tab4:
        st.dataframe(last_matches_df[match_cols], use_container_width=True, hide_index=True)


def main() -> None:
    set_page_config()
    render_intro()

    mode, cups, random_seed, file_bytes = render_sidebar()

    st.markdown("---")
    st.subheader("Interaja com o maior evento esportivo do planeta!")
    st.write(
        "🍀 **Traga sorte para a Seleção!** Aperte o botão e ajude a manifestar o Hexa no nosso servidor. "
    "Se o Brasil não sair campeão na sua simulação, não se desespere: **foi apenas um erro de arredondamento estatístico.** "
    "Tente novamente até a realidade (ou o algoritmo) colaborar!"
    "Simule agora e veja quem levanta a taça."
    )

    if file_bytes is None:
        st.info(
            "Envie a planilha `.xlsx` na barra lateral para habilitar a simulacao. "
            "Se voce colocar `copachaveamento.xlsx` na mesma pasta do app, ele passa a carregar automaticamente."
        )
        return

    if st.button("Simular minha Copa", type="primary", use_container_width=True):
        with st.spinner("Rodando a simulacao da Copa..."):
            try:
                result = run_simulation_from_upload(
                    file_bytes=file_bytes,
                    mode=mode,
                    cups=cups,
                    random_seed=random_seed,
                )
            except Exception as exc:
                st.error(f"Nao foi possivel rodar a simulacao: {exc.__class__.__name__}: {exc}")
                return

        render_results(result, cups)


if __name__ == "__main__":
    main()
