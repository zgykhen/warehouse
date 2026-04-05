import csv
import os
import unicodedata
from typing import Dict, List, Tuple

from app_paths import APP_DIR, BOM_FILENAME, DESCRIPTION_FILENAME


def normalizar_referencia(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return str(valor).strip().upper()


def normalizar_cabecalho_csv(valor) -> str:
    txt = str(valor or "").strip().lower()
    txt = unicodedata.normalize("NFKD", txt)
    return "".join(ch for ch in txt if not unicodedata.combining(ch))


def detetar_delimitador_csv(caminho: str, default: str = ";") -> str:
    try:
        with open(caminho, mode="r", newline="", encoding="utf-8-sig") as f:
            sample = f.read(4096)
    except OSError:
        return default

    if not sample:
        return default

    try:
        return csv.Sniffer().sniff(sample, delimiters=";,\t|").delimiter
    except csv.Error:
        primeira_linha = sample.splitlines()[0] if sample.splitlines() else ""
        if primeira_linha.count(";") >= primeira_linha.count(","):
            return ";"
        if primeira_linha.count(",") > 0:
            return ","
        if "\t" in primeira_linha:
            return "\t"
        if "|" in primeira_linha:
            return "|"
        return default


def normalizar_quantidade_csv(valor, default: int = 1) -> int:
    txt = str(valor).strip() if valor is not None else ""
    if not txt:
        return default
    txt = txt.replace(",", ".")
    try:
        qtd = int(float(txt))
        return qtd if qtd > 0 else default
    except (TypeError, ValueError):
        return default


def carregar_descricoes(description_path: str = "") -> Dict[str, str]:
    """Lê Description.csv e devolve dicionário {REFERENCIA: DESCRIPTION}."""
    descricoes: Dict[str, str] = {}
    if not description_path:
        description_path = os.path.join(APP_DIR, DESCRIPTION_FILENAME)
    if not os.path.isfile(description_path):
        return descricoes

    delimitador = detetar_delimitador_csv(description_path)
    ref_alias = {"reference", "referencia", "ref"}
    desc_alias = {"description", "descricao", "desc"}

    try:
        with open(description_path, mode="r", newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f, delimiter=delimitador)
            fieldnames = reader.fieldnames or []
            field_map = {
                normalizar_cabecalho_csv(nome): nome
                for nome in fieldnames
                if nome is not None
            }

            ref_key = next((field_map[k] for k in ("reference", "referencia", "ref") if k in field_map), None)
            desc_key = next((field_map[k] for k in ("description", "descricao", "desc") if k in field_map), None)

            if ref_key and desc_key:
                for row in reader:
                    ref = normalizar_referencia(row.get(ref_key))
                    if not ref:
                        continue
                    descricoes[ref] = str(row.get(desc_key) or "").strip()
                return descricoes
    except (OSError, csv.Error):
        pass

    # fallback posicional: coluna 1=Reference, coluna 2=Description
    try:
        with open(description_path, mode="r", newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=delimitador)
            for idx, row in enumerate(reader):
                if not row:
                    continue

                if idx == 0 and len(row) >= 2:
                    h0 = normalizar_cabecalho_csv(row[0])
                    h1 = normalizar_cabecalho_csv(row[1])
                    if h0 in ref_alias and h1 in desc_alias:
                        continue

                ref = normalizar_referencia(row[0] if len(row) > 0 else "")
                if not ref:
                    continue

                descricoes[ref] = str(row[1] if len(row) > 1 else "").strip()
    except (OSError, csv.Error):
        pass

    return descricoes


def carregar_lotes_completos(bom_path: str = "") -> Dict[str, List[Tuple[str, int]]]:
    """Lê BOM.csv e devolve lotes no formato {seat: [(component, qty), ...]}."""
    if not bom_path:
        bom_path = os.path.join(APP_DIR, BOM_FILENAME)

    lotes: Dict[str, List[Tuple[str, int]]] = {}
    if not os.path.isfile(bom_path):
        return lotes

    delimitador = detetar_delimitador_csv(bom_path)
    seat_alias = {"seat"}
    comp_alias = {"component", "referencia", "reference", "ref"}
    qty_alias = {"quantity", "quantidade", "qty", "qtd"}

    try:
        with open(bom_path, mode="r", newline="", encoding="utf-8-sig") as f:
            reader = csv.DictReader(f, delimiter=delimitador)
            fieldnames = reader.fieldnames or []
            field_map = {
                normalizar_cabecalho_csv(nome): nome
                for nome in fieldnames
                if nome is not None
            }

            seat_key = next((field_map[k] for k in ("seat",) if k in field_map), None)
            comp_key = next(
                (field_map[k] for k in ("component", "referencia", "reference", "ref") if k in field_map),
                None,
            )
            qty_key = next(
                (field_map[k] for k in ("quantity", "quantidade", "qty", "qtd") if k in field_map),
                None,
            )

            if seat_key and comp_key:
                for row in reader:
                    seat = str(row.get(seat_key) or "").strip()
                    comp = normalizar_referencia(row.get(comp_key))
                    if not seat or not comp:
                        continue

                    qty = normalizar_quantidade_csv(row.get(qty_key), default=1) if qty_key else 1
                    lotes.setdefault(seat, []).append((comp, qty))
                return lotes
    except (OSError, csv.Error):
        pass

    # fallback posicional: Seat;Component;Quantity
    try:
        with open(bom_path, mode="r", newline="", encoding="utf-8-sig") as f:
            reader = csv.reader(f, delimiter=delimitador)
            for idx, row in enumerate(reader):
                if not row:
                    continue

                if idx == 0 and len(row) >= 3:
                    h0 = normalizar_cabecalho_csv(row[0])
                    h1 = normalizar_cabecalho_csv(row[1])
                    h2 = normalizar_cabecalho_csv(row[2])
                    if h0 in seat_alias and h1 in comp_alias and h2 in qty_alias:
                        continue

                seat = str(row[0]).strip() if len(row) > 0 else ""
                comp = normalizar_referencia(row[1] if len(row) > 1 else "")
                if not seat or not comp:
                    continue

                qty = normalizar_quantidade_csv(row[2] if len(row) > 2 else "", default=1)
                lotes.setdefault(seat, []).append((comp, qty))
    except (OSError, csv.Error):
        pass

    return lotes

