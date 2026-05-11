from __future__ import annotations

import re
import secrets
import unicodedata
from datetime import datetime
from pathlib import Path
from shutil import copy2

import numpy as np
import pandas as pd

try:
    from tqdm.auto import tqdm
except ImportError:  # pragma: no cover
    def tqdm(iterable, **kwargs):
        return iterable


DEFAULT_EXCEL_PATH = Path(r"C:\Users\rhena\Downloads\Modelagem - Copa\copachaveamento.xlsx")
DEFAULT_OUTPUT_PATH = DEFAULT_EXCEL_PATH.with_name(
    f"{DEFAULT_EXCEL_PATH.stem}_simulado.xlsx"
)

DEFAULT_SIMULATION_PRESET = "equilibrado"

SIMULATION_PRESETS = {
    "conservador": {
        "probability_concentration": 70.0,
        "probability_tilt": 0.70,
        "attack_weight": 0.58,
        "defense_weight": 0.42,
        "min_goals_lambda": 0.20,
        "max_goals_lambda": 4.20,
    },
    "equilibrado": {
        "probability_concentration": 40.0,
        "probability_tilt": 0.55,
        "attack_weight": 0.60,
        "defense_weight": 0.40,
        "min_goals_lambda": 0.15,
        "max_goals_lambda": 4.75,
    },
    "caotico": {
        "probability_concentration": 18.0,
        "probability_tilt": 0.38,
        "attack_weight": 0.64,
        "defense_weight": 0.36,
        "min_goals_lambda": 0.10,
        "max_goals_lambda": 5.40,
    },
}

DEFAULT_NAME_ALIASES = {
    "Korea Republic": "South Korea",
    "Bosnia and Herzegovina": "Bosnia-Herzegovina",
    "Bosnia Herzegovina": "Bosnia-Herzegovina",
    "Turkiye": "Turkiye",
    "Turkiye ": "Turkiye",
    "Türkiye": "Turkiye",
    "Congo DR": "Congo",
    "DR Congo": "Congo",
}


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def ascii_key(value: object) -> str:
    text = clean_text(value)
    text = unicodedata.normalize("NFKD", text)
    text = text.encode("ascii", "ignore").decode("ascii")
    text = text.lower().replace("&", " and ")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def canonical_name_key(value: object) -> str:
    return ascii_key(value).replace(" ", "")


def resolve_execution_seed(random_seed: int | None) -> int:
    if random_seed is None:
        return secrets.randbits(64)
    return int(random_seed)


def get_simulation_preset(preset_name: str | None = None) -> dict[str, float]:
    key = ascii_key(preset_name) or DEFAULT_SIMULATION_PRESET
    if key not in SIMULATION_PRESETS:
        valid = ", ".join(sorted(SIMULATION_PRESETS))
        raise KeyError(f"Preset invalido: {preset_name!r}. Use um de: {valid}")
    return SIMULATION_PRESETS[key].copy()


def resolve_simulation_parameters(
    *,
    preset_name: str | None,
    probability_concentration: float | None,
    probability_tilt: float | None,
    attack_weight: float | None,
    defense_weight: float | None,
    min_goals_lambda: float | None,
    max_goals_lambda: float | None,
) -> tuple[str, dict[str, float]]:
    resolved_name = ascii_key(preset_name) or DEFAULT_SIMULATION_PRESET
    params = get_simulation_preset(resolved_name)
    overrides = {
        "probability_concentration": probability_concentration,
        "probability_tilt": probability_tilt,
        "attack_weight": attack_weight,
        "defense_weight": defense_weight,
        "min_goals_lambda": min_goals_lambda,
        "max_goals_lambda": max_goals_lambda,
    }
    for key, value in overrides.items():
        if value is not None:
            params[key] = float(value)
    return resolved_name, params


def find_column(columns: list[str], *candidates: str) -> str:
    lookup = {ascii_key(column): column for column in columns}
    for candidate in candidates:
        key = ascii_key(candidate)
        if key in lookup:
            return lookup[key]
    raise KeyError(f"Nao encontrei nenhuma das colunas esperadas: {candidates}")


def parse_match_number(label: object) -> int:
    match = re.search(r"(\d+)", clean_text(label))
    if not match:
        raise ValueError(f"Rotulo de match invalido: {label!r}")
    return int(match.group(1))


def stage_from_match_number(match_no: int) -> str:
    if 1 <= match_no <= 72:
        return "Grupos"
    if 73 <= match_no <= 88:
        return "Round of 32"
    if 89 <= match_no <= 96:
        return "Oitavas"
    if 97 <= match_no <= 100:
        return "Quartas"
    if 101 <= match_no <= 102:
        return "Semifinal"
    if match_no == 103:
        return "3o Lugar"
    if match_no == 104:
        return "Final"
    raise ValueError(f"Match fora do intervalo esperado: {match_no}")


def build_name_lookup(
    selections_df: pd.DataFrame,
    aliases: dict[str, str] | None = None,
) -> dict[str, str]:
    aliases = aliases or {}
    lookup: dict[str, str] = {}

    for row in selections_df.itertuples(index=False):
        lookup[canonical_name_key(row.selecao)] = row.selecao
        lookup[canonical_name_key(row.codigo)] = row.selecao

    for alias, canonical in aliases.items():
        if canonical not in set(selections_df["selecao"]):
            raise KeyError(f"Alias aponta para selecao inexistente: {alias} -> {canonical}")
        lookup[canonical_name_key(alias)] = canonical

    return lookup


def canonicalize_team_name(
    value: object,
    lookup: dict[str, str],
    *,
    allow_missing: bool = False,
) -> str | None:
    key = canonical_name_key(value)
    if key in lookup:
        return lookup[key]
    if allow_missing:
        return None
    raise KeyError(f"Selecao nao encontrada: {value!r}")


def load_selections(excel_path: Path, sheet_name: str = "Seleções") -> pd.DataFrame:
    raw = pd.read_excel(excel_path, sheet_name=sheet_name)
    raw.columns = [clean_text(col) for col in raw.columns]

    col_id = find_column(raw.columns.tolist(), "id")
    col_team = find_column(raw.columns.tolist(), "Seleção", "Selecao", "Team")
    col_code = find_column(raw.columns.tolist(), "Código", "Codigo", "Code")
    col_group = find_column(raw.columns.tolist(), "Grupo", "Group")
    col_players = find_column(raw.columns.tolist(), "Média dos jogadores", "Media dos jogadores")
    col_staff = find_column(raw.columns.tolist(), "Comissão Técnica", "Comissao Tecnica")
    col_fs = find_column(raw.columns.tolist(), "FS_Norm")
    col_er = find_column(raw.columns.tolist(), "ER_norm")

    df = raw[
        [col_id, col_team, col_code, col_group, col_players, col_staff, col_fs, col_er]
    ].copy()
    df.columns = [
        "id",
        "selecao",
        "codigo",
        "grupo",
        "media_jogadores",
        "comissao_tecnica",
        "fs_norm",
        "er_norm",
    ]

    df["id"] = pd.to_numeric(df["id"], errors="coerce")
    df = df[df["id"].notna() & df["selecao"].notna()].copy()
    df["id"] = df["id"].astype(int)

    for column in ["selecao", "codigo", "grupo"]:
        df[column] = df[column].map(clean_text)

    for column in ["media_jogadores", "comissao_tecnica", "fs_norm", "er_norm"]:
        df[column] = pd.to_numeric(df[column], errors="coerce")

    df = df[df["selecao"] != ""].copy()
    df = df.drop_duplicates(subset=["selecao"]).reset_index(drop=True)
    return df


def load_saldo(
    excel_path: Path,
    name_lookup: dict[str, str],
    sheet_name: str = "Saldo",
) -> pd.DataFrame:
    raw = pd.read_excel(excel_path, sheet_name=sheet_name)
    raw.columns = [clean_text(col) for col in raw.columns]

    col_team = find_column(raw.columns.tolist(), "Seleção", "Selecao")
    col_gp_avg = find_column(raw.columns.tolist(), "Média de gols feitos", "Media de gols feitos")
    col_gc_avg = find_column(raw.columns.tolist(), "Média de gols sofridos", "Media de gols sofridos")
    col_points = find_column(raw.columns.tolist(), "Saldo de Gols (SG)", "SG", "Saldo")

    df = raw[[col_team, col_gp_avg, col_gc_avg, col_points]].copy()
    df.columns = ["selecao", "media_gols_feitos", "media_gols_sofridos", "saldo_20j"]
    df = df[df["selecao"].notna()].copy()
    df["selecao"] = df["selecao"].map(lambda x: canonicalize_team_name(x, name_lookup))
    for column in ["media_gols_feitos", "media_gols_sofridos", "saldo_20j"]:
        df[column] = pd.to_numeric(df[column], errors="coerce")
    df = df.drop_duplicates(subset=["selecao"]).reset_index(drop=True)
    return df


def load_probabilities(
    excel_path: Path,
    team_names: list[str],
    name_lookup: dict[str, str],
    sheet_name: str = "Probabilidades",
) -> pd.DataFrame:
    raw = pd.read_excel(excel_path, sheet_name=sheet_name, index_col=0)
    raw.index = [clean_text(idx) for idx in raw.index]
    raw.columns = [clean_text(col) for col in raw.columns]

    result = pd.DataFrame(np.nan, index=team_names, columns=team_names, dtype=float)
    for team in team_names:
        result.at[team, team] = 0.5

    row_map: dict[str, str] = {}
    for label in raw.index:
        canonical = canonicalize_team_name(label, name_lookup, allow_missing=True)
        if canonical is not None:
            row_map[label] = canonical

    col_map: dict[str, str] = {}
    for label in raw.columns:
        canonical = canonicalize_team_name(label, name_lookup, allow_missing=True)
        if canonical is not None:
            col_map[label] = canonical

    for raw_row, team_a in row_map.items():
        for raw_col, team_b in col_map.items():
            if team_a == team_b:
                continue
            value = pd.to_numeric(pd.Series([raw.at[raw_row, raw_col]]), errors="coerce").iloc[0]
            if pd.notna(value):
                result.at[team_a, team_b] = float(value)

    for i, team_a in enumerate(team_names):
        for team_b in team_names[i + 1 :]:
            p_ab = result.at[team_a, team_b]
            p_ba = result.at[team_b, team_a]
            candidates = []
            if pd.notna(p_ab):
                candidates.append(float(p_ab))
            if pd.notna(p_ba):
                candidates.append(1.0 - float(p_ba))
            if candidates:
                probability = float(np.clip(np.mean(candidates), 0.02, 0.98))
            else:
                probability = 0.5
            result.at[team_a, team_b] = probability
            result.at[team_b, team_a] = 1.0 - probability

    return result


def load_chaveamento(
    excel_path: Path,
    selections_df: pd.DataFrame,
    name_lookup: dict[str, str],
    sheet_name: str = "Chaveamento",
) -> tuple[pd.DataFrame, pd.DataFrame]:
    raw = pd.read_excel(excel_path, sheet_name=sheet_name)
    raw.columns = [clean_text(col) for col in raw.columns]

    col_round = find_column(raw.columns.tolist(), "Rodada")
    col_a = find_column(raw.columns.tolist(), "Time A")
    col_b = find_column(raw.columns.tolist(), "Time B")

    df = raw[[col_round, col_a, col_b]].copy()
    df.columns = ["rodada", "slot_a", "slot_b"]
    df = df[df["rodada"].notna()].copy()
    df["match_no"] = df["rodada"].map(parse_match_number)
    df["stage"] = df["match_no"].map(stage_from_match_number)
    df["slot_a"] = df["slot_a"].map(clean_text)
    df["slot_b"] = df["slot_b"].map(clean_text)

    group_map = selections_df.set_index("selecao")["grupo"].to_dict()

    group_matches = df[df["match_no"].between(1, 72)].copy()
    group_matches["team_a"] = group_matches["slot_a"].map(
        lambda value: canonicalize_team_name(value, name_lookup)
    )
    group_matches["team_b"] = group_matches["slot_b"].map(
        lambda value: canonicalize_team_name(value, name_lookup)
    )
    group_matches["grupo"] = group_matches["team_a"].map(group_map)

    knockout_matches = df[df["match_no"].between(73, 104)].copy()
    return (
        group_matches[["match_no", "stage", "grupo", "team_a", "team_b"]].reset_index(drop=True),
        knockout_matches[["match_no", "stage", "slot_a", "slot_b"]].reset_index(drop=True),
    )


def build_team_data(selections_df: pd.DataFrame, saldo_df: pd.DataFrame) -> pd.DataFrame:
    merged = selections_df.merge(saldo_df, on="selecao", how="left", validate="1:1")
    numeric_cols = [
        "media_jogadores",
        "comissao_tecnica",
        "fs_norm",
        "er_norm",
        "media_gols_feitos",
        "media_gols_sofridos",
    ]
    if merged[numeric_cols].isna().any().any():
        missing = merged[merged[numeric_cols].isna().any(axis=1)]["selecao"].tolist()
        raise ValueError(f"Faltam dados numericos para: {missing}")

    weights = {
        "media_jogadores": 0.35,
        "comissao_tecnica": 0.15,
        "fs_norm": 0.15,
        "er_norm": 0.35,
    }

    power = np.zeros(len(merged), dtype=float)
    for column, weight in weights.items():
        series = merged[column].astype(float)
        spread = series.max() - series.min()
        if spread > 0:
            normalized = (series - series.min()) / spread
        else:
            normalized = pd.Series(np.full(len(series), 0.5), index=series.index)
        power += normalized.to_numpy() * weight

    merged["power_score"] = power
    return merged.set_index("selecao").sort_index()


def estimate_probability_from_power(team_a: str, team_b: str, team_data: pd.DataFrame) -> float:
    delta = float(team_data.at[team_a, "power_score"] - team_data.at[team_b, "power_score"])
    return float(np.clip(1.0 / (1.0 + np.exp(-5.0 * delta)), 0.02, 0.98))


def get_base_probability(
    team_a: str,
    team_b: str,
    probability_matrix: pd.DataFrame,
    team_data: pd.DataFrame,
) -> float:
    value = probability_matrix.at[team_a, team_b]
    if pd.isna(value):
        return estimate_probability_from_power(team_a, team_b, team_data)
    return float(np.clip(float(value), 0.02, 0.98))


def sample_match_probability(
    base_probability: float,
    rng: np.random.Generator,
    concentration: float,
) -> float:
    base_probability = float(np.clip(base_probability, 0.02, 0.98))
    if concentration <= 0:
        return base_probability
    alpha = max(base_probability * concentration, 1e-3)
    beta = max((1.0 - base_probability) * concentration, 1e-3)
    return float(rng.beta(alpha, beta))


def compute_goal_rates(
    team_a: str,
    team_b: str,
    sampled_probability: float,
    team_data: pd.DataFrame,
    *,
    attack_weight: float,
    defense_weight: float,
    probability_tilt: float,
    min_lambda: float,
    max_lambda: float,
) -> tuple[float, float]:
    attack_a = float(team_data.at[team_a, "media_gols_feitos"])
    defense_a = float(team_data.at[team_a, "media_gols_sofridos"])
    attack_b = float(team_data.at[team_b, "media_gols_feitos"])
    defense_b = float(team_data.at[team_b, "media_gols_sofridos"])

    base_a = max(attack_weight * attack_a + defense_weight * defense_b, min_lambda)
    base_b = max(attack_weight * attack_b + defense_weight * defense_a, min_lambda)

    total_goals = max(base_a + base_b, 2.0 * min_lambda)
    sampled_probability = float(np.clip(sampled_probability, 1e-4, 1.0 - 1e-4))
    logit = np.log(sampled_probability / (1.0 - sampled_probability))
    tilt = np.exp(probability_tilt * logit)

    adjusted_a = base_a * tilt
    adjusted_b = base_b / tilt
    scale = total_goals / max(adjusted_a + adjusted_b, 1e-9)

    lambda_a = float(np.clip(adjusted_a * scale, min_lambda, max_lambda))
    lambda_b = float(np.clip(adjusted_b * scale, min_lambda, max_lambda))
    return lambda_a, lambda_b


def simulate_match(
    team_a: str,
    team_b: str,
    team_data: pd.DataFrame,
    probability_matrix: pd.DataFrame,
    rng: np.random.Generator,
    *,
    allow_draw: bool,
    probability_concentration: float,
    probability_tilt: float,
    attack_weight: float,
    defense_weight: float,
    min_lambda: float,
    max_lambda: float,
) -> dict[str, object]:
    base_probability = get_base_probability(team_a, team_b, probability_matrix, team_data)
    sampled_probability = sample_match_probability(base_probability, rng, probability_concentration)
    lambda_a, lambda_b = compute_goal_rates(
        team_a,
        team_b,
        sampled_probability,
        team_data,
        attack_weight=attack_weight,
        defense_weight=defense_weight,
        probability_tilt=probability_tilt,
        min_lambda=min_lambda,
        max_lambda=max_lambda,
    )

    goals_a = int(rng.poisson(lambda_a))
    goals_b = int(rng.poisson(lambda_b))

    winner = ""
    loser = ""
    decision = "Empate"

    if goals_a > goals_b:
        winner, loser, decision = team_a, team_b, "Tempo normal"
    elif goals_b > goals_a:
        winner, loser, decision = team_b, team_a, "Tempo normal"
    elif not allow_draw:
        if rng.random() < sampled_probability:
            winner, loser = team_a, team_b
        else:
            winner, loser = team_b, team_a
        decision = "Penaltis"

    return {
        "team_a": team_a,
        "team_b": team_b,
        "goals_a": goals_a,
        "goals_b": goals_b,
        "winner": winner,
        "loser": loser,
        "decision": decision,
        "base_probability_a": base_probability,
        "sampled_probability_a": sampled_probability,
        "lambda_a": lambda_a,
        "lambda_b": lambda_b,
    }


def init_group_tables(group_members: dict[str, list[str]]) -> dict[str, pd.DataFrame]:
    tables = {}
    for group, teams in group_members.items():
        tables[group] = pd.DataFrame(
            0,
            index=teams,
            columns=["P", "J", "V", "E", "D", "GP", "GC", "SG"],
            dtype=int,
        )
    return tables


def update_group_table(table: pd.DataFrame, team_a: str, team_b: str, goals_a: int, goals_b: int) -> None:
    table.loc[team_a, ["J", "GP", "GC"]] += [1, goals_a, goals_b]
    table.loc[team_b, ["J", "GP", "GC"]] += [1, goals_b, goals_a]
    table.loc[team_a, "SG"] = table.loc[team_a, "GP"] - table.loc[team_a, "GC"]
    table.loc[team_b, "SG"] = table.loc[team_b, "GP"] - table.loc[team_b, "GC"]
    if goals_a > goals_b:
        table.loc[team_a, ["V", "P"]] += [1, 3]
        table.loc[team_b, "D"] += 1
    elif goals_b > goals_a:
        table.loc[team_b, ["V", "P"]] += [1, 3]
        table.loc[team_a, "D"] += 1
    else:
        table.loc[team_a, ["E", "P"]] += [1, 1]
        table.loc[team_b, ["E", "P"]] += [1, 1]


def rank_group_table(
    table: pd.DataFrame,
    team_data: pd.DataFrame,
    rng: np.random.Generator,
) -> pd.DataFrame:
    ranked = table.copy()
    ranked["selecao"] = ranked.index
    ranked["power_score"] = ranked["selecao"].map(team_data["power_score"])
    ranked["sorteio"] = rng.random(len(ranked))
    ranked = ranked.sort_values(
        by=["P", "SG", "GP", "V", "power_score", "sorteio"],
        ascending=[False, False, False, False, False, False],
    ).reset_index(drop=True)
    ranked.insert(0, "classificacao", np.arange(1, len(ranked) + 1))
    return ranked[
        ["classificacao", "selecao", "P", "J", "V", "E", "D", "GP", "GC", "SG"]
    ]


def select_best_thirds(
    group_rankings: dict[str, pd.DataFrame],
    team_data: pd.DataFrame,
    rng: np.random.Generator,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    rows = []
    for group, ranking in group_rankings.items():
        third = ranking.iloc[2].copy()
        third["grupo"] = group
        rows.append(third)
    thirds = pd.DataFrame(rows)
    thirds["power_score"] = thirds["selecao"].map(team_data["power_score"])
    thirds["sorteio"] = rng.random(len(thirds))
    thirds = thirds.sort_values(
        by=["P", "SG", "GP", "V", "power_score", "sorteio"],
        ascending=[False, False, False, False, False, False],
    ).reset_index(drop=True)
    thirds.insert(0, "ranking_3os", np.arange(1, len(thirds) + 1))
    thirds["classifica"] = thirds["ranking_3os"] <= 8
    qualified = thirds[thirds["classifica"]].copy().reset_index(drop=True)
    return thirds, qualified


def parse_knockout_slot(slot_text: str) -> dict[str, object]:
    text = clean_text(slot_text)
    low = ascii_key(text)

    third = re.match(r"^group ([a-l](?: [a-l])*) third place$", low)
    if third:
        groups = tuple(part.upper() for part in third.group(1).split())
        return {"kind": "third_place", "groups": groups}

    group_pos = re.match(r"^group ([a-l]) (winner|winners|runner up|runners up)$", low)
    if group_pos:
        return {
            "kind": "group_position",
            "group": group_pos.group(1).upper(),
            "position": "winner" if "winner" in group_pos.group(2) else "runner_up",
        }

    winner_match = re.match(r"^winner match (\d+)$", low)
    if winner_match:
        return {"kind": "match_winner", "match_no": int(winner_match.group(1))}

    loser_match = re.match(r"^runner up match (\d+)$", low)
    if loser_match:
        return {"kind": "match_loser", "match_no": int(loser_match.group(1))}

    return {"kind": "team", "team": text}


def build_third_place_slots(knockout_matches: pd.DataFrame) -> list[dict[str, object]]:
    slots = []
    for row in knockout_matches.sort_values("match_no").itertuples(index=False):
        for side, slot in (("A", row.slot_a), ("B", row.slot_b)):
            descriptor = parse_knockout_slot(slot)
            if descriptor["kind"] == "third_place":
                slots.append(
                    {
                        "slot_id": f"{row.match_no}_{side}",
                        "match_no": row.match_no,
                        "side": side,
                        "allowed_groups": descriptor["groups"],
                    }
                )
    return slots


def assign_best_thirds_to_slots(
    best_thirds_df: pd.DataFrame,
    third_slots: list[dict[str, object]],
) -> dict[str, str]:
    ordered_slots = sorted(third_slots, key=lambda slot: (slot["match_no"], slot["side"]))
    teams = best_thirds_df[["selecao", "grupo"]].to_dict("records")

    def backtrack(index: int, remaining: list[dict[str, str]]) -> dict[str, str] | None:
        if index >= len(ordered_slots):
            return {}
        slot = ordered_slots[index]
        candidates = [team for team in remaining if team["grupo"] in slot["allowed_groups"]]
        for candidate in candidates:
            next_remaining = [team for team in remaining if team["selecao"] != candidate["selecao"]]
            result = backtrack(index + 1, next_remaining)
            if result is not None:
                result[slot["slot_id"]] = candidate["selecao"]
                return result
        return None

    assignment = backtrack(0, teams)
    if assignment is None:
        raise ValueError("Nao foi possivel alocar os terceiros colocados.")
    return assignment


def resolve_knockout_participant(
    slot_text: str,
    match_no: int,
    side: str,
    context: dict[str, object],
    name_lookup: dict[str, str],
) -> str:
    descriptor = parse_knockout_slot(slot_text)
    kind = descriptor["kind"]
    if kind == "team":
        return canonicalize_team_name(descriptor["team"], name_lookup)
    if kind == "group_position":
        ranking = context["group_rankings"][descriptor["group"]]
        idx = 0 if descriptor["position"] == "winner" else 1
        return str(ranking.iloc[idx]["selecao"])
    if kind == "third_place":
        return context["third_assignment"][f"{match_no}_{side}"]
    if kind == "match_winner":
        return context["match_winners"][descriptor["match_no"]]
    if kind == "match_loser":
        return context["match_losers"][descriptor["match_no"]]
    raise ValueError(f"Referencia invalida: {slot_text}")


def simulate_group_stage(
    group_matches: pd.DataFrame,
    group_members: dict[str, list[str]],
    team_data: pd.DataFrame,
    probability_matrix: pd.DataFrame,
    rng: np.random.Generator,
    *,
    probability_concentration: float,
    probability_tilt: float,
    attack_weight: float,
    defense_weight: float,
    min_lambda: float,
    max_lambda: float,
) -> tuple[dict[str, pd.DataFrame], pd.DataFrame, pd.DataFrame, list[dict[str, object]]]:
    tables = init_group_tables(group_members)
    logs: list[dict[str, object]] = []
    for row in group_matches.sort_values("match_no").itertuples(index=False):
        result = simulate_match(
            row.team_a,
            row.team_b,
            team_data,
            probability_matrix,
            rng,
            allow_draw=True,
            probability_concentration=probability_concentration,
            probability_tilt=probability_tilt,
            attack_weight=attack_weight,
            defense_weight=defense_weight,
            min_lambda=min_lambda,
            max_lambda=max_lambda,
        )
        update_group_table(tables[row.grupo], row.team_a, row.team_b, result["goals_a"], result["goals_b"])
        logs.append(
            {
                "match_no": row.match_no,
                "fase": row.stage,
                "grupo": row.grupo,
                "referencia_a": row.team_a,
                "referencia_b": row.team_b,
                "time_a": row.team_a,
                "time_b": row.team_b,
                "gols_a": result["goals_a"],
                "gols_b": result["goals_b"],
                "vencedor": result["winner"],
                "perdedor": result["loser"],
                "decisao": result["decision"],
                "prob_base_a": result["base_probability_a"],
                "prob_sorteada_a": result["sampled_probability_a"],
                "lambda_a": result["lambda_a"],
                "lambda_b": result["lambda_b"],
            }
        )

    rankings = {group: rank_group_table(table, team_data, rng) for group, table in tables.items()}
    all_thirds, best_thirds = select_best_thirds(rankings, team_data, rng)
    return rankings, all_thirds, best_thirds, logs


def simulate_knockout(
    knockout_matches: pd.DataFrame,
    group_rankings: dict[str, pd.DataFrame],
    best_thirds_df: pd.DataFrame,
    team_data: pd.DataFrame,
    probability_matrix: pd.DataFrame,
    name_lookup: dict[str, str],
    rng: np.random.Generator,
    *,
    probability_concentration: float,
    probability_tilt: float,
    attack_weight: float,
    defense_weight: float,
    min_lambda: float,
    max_lambda: float,
) -> dict[str, object]:
    third_assignment = assign_best_thirds_to_slots(best_thirds_df, build_third_place_slots(knockout_matches))
    context: dict[str, object] = {
        "group_rankings": group_rankings,
        "third_assignment": third_assignment,
        "match_winners": {},
        "match_losers": {},
    }
    logs: list[dict[str, object]] = []

    for row in knockout_matches.sort_values("match_no").itertuples(index=False):
        team_a = resolve_knockout_participant(row.slot_a, row.match_no, "A", context, name_lookup)
        team_b = resolve_knockout_participant(row.slot_b, row.match_no, "B", context, name_lookup)
        result = simulate_match(
            team_a,
            team_b,
            team_data,
            probability_matrix,
            rng,
            allow_draw=False,
            probability_concentration=probability_concentration,
            probability_tilt=probability_tilt,
            attack_weight=attack_weight,
            defense_weight=defense_weight,
            min_lambda=min_lambda,
            max_lambda=max_lambda,
        )
        context["match_winners"][row.match_no] = result["winner"]
        context["match_losers"][row.match_no] = result["loser"]
        logs.append(
            {
                "match_no": row.match_no,
                "fase": row.stage,
                "grupo": "",
                "referencia_a": row.slot_a,
                "referencia_b": row.slot_b,
                "time_a": team_a,
                "time_b": team_b,
                "gols_a": result["goals_a"],
                "gols_b": result["goals_b"],
                "vencedor": result["winner"],
                "perdedor": result["loser"],
                "decisao": result["decision"],
                "prob_base_a": result["base_probability_a"],
                "prob_sorteada_a": result["sampled_probability_a"],
                "lambda_a": result["lambda_a"],
                "lambda_b": result["lambda_b"],
            }
        )

    context["match_logs"] = logs
    return context


def build_tournament_stats(simulation: dict[str, object], team_data: pd.DataFrame) -> pd.DataFrame:
    stats = pd.DataFrame(
        0,
        index=team_data.index.tolist(),
        columns=["Pontos", "J", "V", "E", "D", "GP", "GC", "SG"],
        dtype=int,
    )
    for log in simulation["group_match_logs"] + simulation["knockout_match_logs"]:
        team_a = log["time_a"]
        team_b = log["time_b"]
        goals_a = int(log["gols_a"])
        goals_b = int(log["gols_b"])
        group_stage = log["fase"] == "Grupos"

        stats.loc[team_a, ["J", "GP", "GC"]] += [1, goals_a, goals_b]
        stats.loc[team_b, ["J", "GP", "GC"]] += [1, goals_b, goals_a]

        if goals_a > goals_b:
            stats.loc[team_a, "V"] += 1
            stats.loc[team_b, "D"] += 1
            if group_stage:
                stats.loc[team_a, "Pontos"] += 3
        elif goals_b > goals_a:
            stats.loc[team_b, "V"] += 1
            stats.loc[team_a, "D"] += 1
            if group_stage:
                stats.loc[team_b, "Pontos"] += 3
        elif group_stage:
            stats.loc[team_a, ["E", "Pontos"]] += [1, 1]
            stats.loc[team_b, ["E", "Pontos"]] += [1, 1]
        else:
            stats.loc[log["vencedor"], "V"] += 1
            stats.loc[log["perdedor"], "D"] += 1

    stats["SG"] = stats["GP"] - stats["GC"]
    return stats


def rank_position_bucket(
    teams: list[str] | set[str],
    tournament_stats: pd.DataFrame,
    team_data: pd.DataFrame,
    rng: np.random.Generator,
    stage_label: str,
) -> pd.DataFrame:
    team_list = [team for team in teams if team]
    if not team_list:
        return pd.DataFrame(
            columns=["selecao", "codigo", "grupo", "etapa_final", "Pontos", "J", "V", "E", "D", "GP", "GC", "SG"]
        )

    frame = tournament_stats.loc[team_list].copy()
    frame["selecao"] = frame.index
    frame["codigo"] = frame["selecao"].map(team_data["codigo"])
    frame["grupo"] = frame["selecao"].map(team_data["grupo"])
    frame["etapa_final"] = stage_label
    frame["power_score"] = frame["selecao"].map(team_data["power_score"])
    frame["sorteio"] = rng.random(len(frame))
    frame = frame.sort_values(
        by=["Pontos", "SG", "GP", "V", "power_score", "sorteio", "selecao"],
        ascending=[False, False, False, False, False, False, True],
    )
    return frame[
        ["selecao", "codigo", "grupo", "etapa_final", "Pontos", "J", "V", "E", "D", "GP", "GC", "SG"]
    ].reset_index(drop=True)


def build_final_ranking(
    simulation: dict[str, object],
    team_data: pd.DataFrame,
    rng: np.random.Generator,
) -> pd.DataFrame:
    team_names = team_data.index.tolist()
    stats = build_tournament_stats(simulation, team_data)

    qualified = set(simulation["qualified"])
    top16 = set(simulation["top16"])
    quarterfinalists = set(simulation["quarterfinalists"])
    semifinalists = set(simulation["semifinalists"])

    group_exit = set(team_names) - qualified
    round32_exit = qualified - top16
    round16_exit = top16 - quarterfinalists
    quarterfinal_exit = quarterfinalists - semifinalists

    frames = [
        rank_position_bucket([simulation["champion"]], stats, team_data, rng, "Campeao"),
        rank_position_bucket([simulation["runner_up"]], stats, team_data, rng, "Vice"),
        rank_position_bucket([simulation["third_place"]], stats, team_data, rng, "3o Lugar"),
        rank_position_bucket([simulation["fourth_place"]], stats, team_data, rng, "4o Lugar"),
        rank_position_bucket(quarterfinal_exit, stats, team_data, rng, "Eliminado nas Quartas"),
        rank_position_bucket(round16_exit, stats, team_data, rng, "Eliminado nas Oitavas"),
        rank_position_bucket(round32_exit, stats, team_data, rng, "Eliminado no Round of 32"),
        rank_position_bucket(group_exit, stats, team_data, rng, "Eliminado nos Grupos"),
    ]

    final_ranking = pd.concat(frames, ignore_index=True)
    final_ranking.insert(0, "posicao", np.arange(1, len(final_ranking) + 1))
    return final_ranking


def simulate_tournament(
    group_matches: pd.DataFrame,
    knockout_matches: pd.DataFrame,
    group_members: dict[str, list[str]],
    team_data: pd.DataFrame,
    probability_matrix: pd.DataFrame,
    name_lookup: dict[str, str],
    rng: np.random.Generator,
    *,
    probability_concentration: float,
    probability_tilt: float,
    attack_weight: float,
    defense_weight: float,
    min_lambda: float,
    max_lambda: float,
) -> dict[str, object]:
    group_rankings, all_thirds, best_thirds, group_logs = simulate_group_stage(
        group_matches,
        group_members,
        team_data,
        probability_matrix,
        rng,
        probability_concentration=probability_concentration,
        probability_tilt=probability_tilt,
        attack_weight=attack_weight,
        defense_weight=defense_weight,
        min_lambda=min_lambda,
        max_lambda=max_lambda,
    )

    knockout = simulate_knockout(
        knockout_matches,
        group_rankings,
        best_thirds,
        team_data,
        probability_matrix,
        name_lookup,
        rng,
        probability_concentration=probability_concentration,
        probability_tilt=probability_tilt,
        attack_weight=attack_weight,
        defense_weight=defense_weight,
        min_lambda=min_lambda,
        max_lambda=max_lambda,
    )

    qualified = set(best_thirds["selecao"].tolist())
    for ranking in group_rankings.values():
        qualified.add(str(ranking.iloc[0]["selecao"]))
        qualified.add(str(ranking.iloc[1]["selecao"]))

    winners = knockout["match_winners"]
    losers = knockout["match_losers"]

    simulation = {
        "qualified": qualified,
        "top16": {winners[m] for m in range(73, 89)},
        "quarterfinalists": {winners[m] for m in range(89, 97)},
        "semifinalists": {winners[m] for m in range(97, 101)},
        "finalists": {winners[101], winners[102]},
        "champion": winners[104],
        "runner_up": losers[104],
        "third_place": winners[103],
        "fourth_place": losers[103],
        "group_rankings": group_rankings,
        "all_thirds": all_thirds,
        "best_thirds": best_thirds,
        "group_match_logs": group_logs,
        "knockout_match_logs": knockout["match_logs"],
        "third_assignment": knockout["third_assignment"],
    }
    simulation["final_ranking"] = build_final_ranking(simulation, team_data, rng)
    return simulation


def update_summary_counts(summary_counts: pd.DataFrame, simulation: dict[str, object], team_names: list[str]) -> None:
    all_teams = set(team_names)
    qualified = set(simulation["qualified"])
    top16 = set(simulation["top16"])
    quarterfinalists = set(simulation["quarterfinalists"])
    semifinalists = set(simulation["semifinalists"])
    finalists = set(simulation["finalists"])

    for team in all_teams - qualified:
        summary_counts.at[team, "Elim_Grupos"] += 1
    for team in qualified:
        summary_counts.at[team, "Mata_Mata"] += 1
    for team in top16:
        summary_counts.at[team, "Oitavas"] += 1
    for team in quarterfinalists:
        summary_counts.at[team, "Quartas"] += 1
    for team in semifinalists:
        summary_counts.at[team, "Semifinal"] += 1
    for team in finalists:
        summary_counts.at[team, "Finalista"] += 1

    summary_counts.at[simulation["champion"], "Campeao"] += 1
    summary_counts.at[simulation["runner_up"], "Vice"] += 1
    summary_counts.at[simulation["third_place"], "Terceiro"] += 1
    summary_counts.at[simulation["fourth_place"], "Quarto"] += 1


def update_position_counts(position_counts: pd.DataFrame, final_ranking: pd.DataFrame) -> None:
    for team_name, position in final_ranking[["selecao", "posicao"]].itertuples(index=False, name=None):
        position_counts.at[team_name, f"Pos_{int(position):02d}"] += 1


def build_ranking_table(
    summary_counts: pd.DataFrame,
    position_counts: pd.DataFrame,
    team_data: pd.DataFrame,
    n_simulations: int,
) -> pd.DataFrame:
    ranking = team_data.reset_index()[["selecao", "codigo", "grupo"]].copy()
    counts_matrix = position_counts.loc[ranking["selecao"]].to_numpy(dtype=float)
    weights = np.arange(1, counts_matrix.shape[1] + 1, dtype=float)
    ranking["Posicao_Media"] = counts_matrix @ weights / float(n_simulations)
    ranking["Posicao_Mais_Provavel"] = counts_matrix.argmax(axis=1) + 1

    mapping = {
        "Elim_Grupos": "Pct_Elim_Grupos",
        "Mata_Mata": "Pct_Mata_Mata",
        "Oitavas": "Pct_Oitavas",
        "Quartas": "Pct_Quartas",
        "Semifinal": "Pct_Semifinal",
        "Finalista": "Pct_Finalista",
        "Campeao": "Pct_Campeao",
        "Vice": "Pct_Vice",
        "Terceiro": "Pct_3o_Lugar",
        "Quarto": "Pct_4o_Lugar",
    }
    for raw_col, out_col in mapping.items():
        ranking[out_col] = summary_counts.loc[ranking["selecao"], raw_col].to_numpy() * 100.0 / n_simulations
    for column in position_counts.columns:
        ranking[column] = position_counts.loc[ranking["selecao"], column].to_numpy() * 100.0 / n_simulations

    ranking = ranking.sort_values(
        by=["Posicao_Media", "Pos_01", "Pos_02", "Pct_Campeao"],
        ascending=[True, False, False, False],
    ).reset_index(drop=True)
    ranking.insert(0, "Ranking_Oficial", np.arange(1, len(ranking) + 1))
    return ranking.round(2)


def match_logs_to_dataframe(simulation: dict[str, object]) -> pd.DataFrame:
    df = pd.DataFrame(simulation["group_match_logs"] + simulation["knockout_match_logs"])
    df = df.sort_values("match_no").reset_index(drop=True)
    for column in ["prob_base_a", "prob_sorteada_a", "lambda_a", "lambda_b"]:
        df[column] = df[column].astype(float).round(4)
    return df


def group_tables_to_dataframe(simulation: dict[str, object]) -> pd.DataFrame:
    qualified_thirds = set(simulation["best_thirds"]["selecao"].tolist())
    frames = []
    for group in sorted(simulation["group_rankings"]):
        frame = simulation["group_rankings"][group].copy()
        frame["grupo"] = group
        frame["qualificado"] = frame.apply(
            lambda row: "Sim"
            if row["classificacao"] <= 2 or (row["classificacao"] == 3 and row["selecao"] in qualified_thirds)
            else "Nao",
            axis=1,
        )
        frames.append(frame)
    return pd.concat(frames, ignore_index=True)


def final_ranking_to_dataframe(simulation: dict[str, object]) -> pd.DataFrame:
    return simulation["final_ranking"].copy()


def build_execution_metadata(
    *,
    execution_seed: int,
    execution_timestamp: str,
    preset_name: str,
    n_simulations: int,
    probability_concentration: float,
    probability_tilt: float,
    attack_weight: float,
    defense_weight: float,
    min_goals_lambda: float,
    max_goals_lambda: float,
    source_path: Path,
    output_path: Path,
) -> pd.DataFrame:
    return pd.DataFrame(
        {
            "Parametro": [
                "execution_seed",
                "execution_timestamp",
                "preset_name",
                "n_simulations",
                "probability_concentration",
                "probability_tilt",
                "attack_weight",
                "defense_weight",
                "min_goals_lambda",
                "max_goals_lambda",
                "source_path",
                "output_path",
            ],
            "Valor": [
                str(execution_seed),
                execution_timestamp,
                preset_name,
                str(n_simulations),
                str(probability_concentration),
                str(probability_tilt),
                str(attack_weight),
                str(defense_weight),
                str(min_goals_lambda),
                str(max_goals_lambda),
                str(source_path),
                str(output_path),
            ],
        }
    )


def export_results(
    ranking_df: pd.DataFrame,
    last_matches_df: pd.DataFrame,
    last_groups_df: pd.DataFrame,
    last_final_ranking_df: pd.DataFrame,
    metadata_df: pd.DataFrame,
    source_path: Path,
    output_path: Path,
    *,
    ranking_sheet: str,
    last_matches_sheet: str,
    last_groups_sheet: str,
    last_final_ranking_sheet: str,
    metadata_sheet: str,
    include_source_sheets: bool = True,
) -> Path:
    if include_source_sheets:
        if output_path.resolve() != source_path.resolve():
            copy2(source_path, output_path)
        with pd.ExcelWriter(
            output_path,
            engine="openpyxl",
            mode="a",
            if_sheet_exists="replace",
        ) as writer:
            ranking_df.to_excel(writer, sheet_name=ranking_sheet, index=False)
            last_matches_df.to_excel(writer, sheet_name=last_matches_sheet, index=False)
            last_groups_df.to_excel(writer, sheet_name=last_groups_sheet, index=False)
            last_final_ranking_df.to_excel(writer, sheet_name=last_final_ranking_sheet, index=False)
            metadata_df.to_excel(writer, sheet_name=metadata_sheet, index=False)
    else:
        if output_path.resolve() == source_path.resolve():
            raise ValueError(
                "Para exportar apenas abas de resultado, output_path deve ser diferente do arquivo base."
            )
        with pd.ExcelWriter(output_path, engine="openpyxl", mode="w") as writer:
            ranking_df.to_excel(writer, sheet_name=ranking_sheet, index=False)
            last_matches_df.to_excel(writer, sheet_name=last_matches_sheet, index=False)
            last_groups_df.to_excel(writer, sheet_name=last_groups_sheet, index=False)
            last_final_ranking_df.to_excel(writer, sheet_name=last_final_ranking_sheet, index=False)
            metadata_df.to_excel(writer, sheet_name=metadata_sheet, index=False)
    return output_path


def run_world_cup_simulation(
    excel_path: Path | str = DEFAULT_EXCEL_PATH,
    *,
    output_path: Path | str | None = None,
    n_simulations: int = 1_000,
    random_seed: int | None = None,
    preset_name: str = DEFAULT_SIMULATION_PRESET,
    export_to_source_workbook: bool = False,
    selection_sheet: str = "Seleções",
    saldo_sheet: str = "Saldo",
    probability_sheet: str = "Probabilidades",
    bracket_sheet: str = "Chaveamento",
    ranking_sheet: str = "Ranking_Oficial",
    last_matches_sheet: str = "Ultima_Simulacao_Jogos",
    last_groups_sheet: str = "Ultima_Simulacao_Grupos",
    last_final_ranking_sheet: str = "Ultima_Simulacao_Ranking",
    metadata_sheet: str = "Metadados_Execucao",
    probability_concentration: float | None = None,
    probability_tilt: float | None = None,
    attack_weight: float | None = None,
    defense_weight: float | None = None,
    min_goals_lambda: float | None = None,
    max_goals_lambda: float | None = None,
    name_aliases: dict[str, str] | None = None,
    include_source_sheets: bool = True,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame, Path]:
    excel_path = Path(excel_path)
    if not excel_path.exists():
        raise FileNotFoundError(f"Arquivo nao encontrado: {excel_path}")

    if output_path is None:
        output_path = excel_path if export_to_source_workbook else excel_path.with_name(
            f"{excel_path.stem}_simulado.xlsx"
        )
    output_path = Path(output_path)

    resolved_preset_name, simulation_params = resolve_simulation_parameters(
        preset_name=preset_name,
        probability_concentration=probability_concentration,
        probability_tilt=probability_tilt,
        attack_weight=attack_weight,
        defense_weight=defense_weight,
        min_goals_lambda=min_goals_lambda,
        max_goals_lambda=max_goals_lambda,
    )

    selections_df = load_selections(excel_path, sheet_name=selection_sheet)
    name_lookup = build_name_lookup(selections_df, name_aliases or DEFAULT_NAME_ALIASES)
    saldo_df = load_saldo(excel_path, name_lookup, sheet_name=saldo_sheet)
    team_data = build_team_data(selections_df, saldo_df)
    probability_matrix = load_probabilities(
        excel_path,
        team_data.index.tolist(),
        name_lookup,
        sheet_name=probability_sheet,
    )
    group_matches, knockout_matches = load_chaveamento(
        excel_path,
        selections_df,
        name_lookup,
        sheet_name=bracket_sheet,
    )
    group_members = selections_df.groupby("grupo", sort=True)["selecao"].apply(list).to_dict()

    team_names = team_data.index.tolist()
    summary_counts = pd.DataFrame(
        0,
        index=team_names,
        columns=["Elim_Grupos", "Mata_Mata", "Oitavas", "Quartas", "Semifinal", "Finalista", "Campeao", "Vice", "Terceiro", "Quarto"],
        dtype=int,
    )
    position_counts = pd.DataFrame(
        0,
        index=team_names,
        columns=[f"Pos_{position:02d}" for position in range(1, len(team_names) + 1)],
        dtype=int,
    )

    execution_seed = resolve_execution_seed(random_seed)
    execution_timestamp = datetime.now().astimezone().isoformat(timespec="seconds")
    rng = np.random.default_rng(execution_seed)

    last_simulation: dict[str, object] | None = None
    for _ in tqdm(range(n_simulations), desc="Simulando Copa 2026"):
        simulation = simulate_tournament(
            group_matches,
            knockout_matches,
            group_members,
            team_data,
            probability_matrix,
            name_lookup,
            rng,
            probability_concentration=simulation_params["probability_concentration"],
            probability_tilt=simulation_params["probability_tilt"],
            attack_weight=simulation_params["attack_weight"],
            defense_weight=simulation_params["defense_weight"],
            min_lambda=simulation_params["min_goals_lambda"],
            max_lambda=simulation_params["max_goals_lambda"],
        )
        update_summary_counts(summary_counts, simulation, team_names)
        update_position_counts(position_counts, simulation["final_ranking"])
        last_simulation = simulation

    if last_simulation is None:
        raise RuntimeError("Nenhuma simulacao foi executada.")

    ranking_df = build_ranking_table(summary_counts, position_counts, team_data, n_simulations)
    last_matches_df = match_logs_to_dataframe(last_simulation)
    last_groups_df = group_tables_to_dataframe(last_simulation)
    last_final_ranking_df = final_ranking_to_dataframe(last_simulation)
    metadata_df = build_execution_metadata(
        execution_seed=execution_seed,
        execution_timestamp=execution_timestamp,
        preset_name=resolved_preset_name,
        n_simulations=n_simulations,
        probability_concentration=simulation_params["probability_concentration"],
        probability_tilt=simulation_params["probability_tilt"],
        attack_weight=simulation_params["attack_weight"],
        defense_weight=simulation_params["defense_weight"],
        min_goals_lambda=simulation_params["min_goals_lambda"],
        max_goals_lambda=simulation_params["max_goals_lambda"],
        source_path=excel_path,
        output_path=output_path,
    )

    for dataframe in (ranking_df, last_matches_df, last_groups_df, last_final_ranking_df):
        dataframe.attrs["execution_seed"] = execution_seed
        dataframe.attrs["execution_timestamp"] = execution_timestamp
        dataframe.attrs["preset_name"] = resolved_preset_name
        dataframe.attrs["simulation_params"] = simulation_params.copy()

    exported_path = export_results(
        ranking_df,
        last_matches_df,
        last_groups_df,
        last_final_ranking_df,
        metadata_df,
        excel_path,
        output_path,
        ranking_sheet=ranking_sheet,
        last_matches_sheet=last_matches_sheet,
        last_groups_sheet=last_groups_sheet,
        last_final_ranking_sheet=last_final_ranking_sheet,
        metadata_sheet=metadata_sheet,
        include_source_sheets=include_source_sheets,
    )

    return ranking_df, last_matches_df, last_groups_df, last_final_ranking_df, exported_path


if __name__ == "__main__":
    ranking, matches, groups, final_ranking, path = run_world_cup_simulation(
        excel_path=DEFAULT_EXCEL_PATH,
        output_path=DEFAULT_OUTPUT_PATH,
        n_simulations=1_000,
        random_seed=None,
        preset_name=DEFAULT_SIMULATION_PRESET,
        export_to_source_workbook=False,
    )
    print(ranking.head(10).to_string(index=False))
    print()
    print(final_ranking.head(10).to_string(index=False))
    print()
    print(f"Preset da execucao: {ranking.attrs.get('preset_name')}")
    print(f"Seed da execucao: {ranking.attrs.get('execution_seed')}")
    print(f"Horario da execucao: {ranking.attrs.get('execution_timestamp')}")
    print()
    print(f"Arquivo exportado: {path}")
