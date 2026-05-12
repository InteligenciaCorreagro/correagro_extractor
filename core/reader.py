"""
Parser del archivo Excel INFORME_EXTRACTOS_YYYY.xlsx
Lee las hojas 'OP VIGENTES' y 'FISICOS COMPRAS', construye modelos de datos.
"""
import openpyxl
from datetime import datetime
import re
from typing import Dict, Tuple, List
from collections import defaultdict

from core.models import (
    InformeData, ClienteOpVigente, RowOpVigente,
    ClienteFisicoCompra, RowFisicoCompra,
    CarteraData, ClienteCartera, RowCartera,
)
from core.pivot_cache_reader import extract_fisicos_from_pivot_cache  # ← CAMBIO 1

SHEET_OP_VIGENTES     = "OP VIGENTES"
SHEET_FISICOS_COMPRAS = "FISICOS COMPRAS"
SHEET_BD              = "BD"


# ──────────────────────────────────────────────────────────────
# Helpers
# ──────────────────────────────────────────────────────────────
def _norm_nit(val) -> str:
    if val is None:
        return ""
    if isinstance(val, bool):
        return ""
    if isinstance(val, int):
        return str(val) if val >= 0 else ""
    if isinstance(val, float):
        if val < 0:
            return ""
        r = round(val)
        if abs(val - r) < 1e-9:
            return str(int(r))
        return ""
    s = str(val).strip()
    return "".join(ch for ch in s if ch.isdigit())


def _looks_like_yyyymmdd_digits(digits: str) -> bool:
    """Evita confundir fechas compactas (YYYYMMDD) con NIT al escanear celdas del encabezado."""
    if len(digits) != 8:
        return False
    try:
        y = int(digits[0:4])
        m = int(digits[4:6])
        d = int(digits[6:8])
    except ValueError:
        return False
    if not (1990 <= y <= 2100):
        return False
    if not (1 <= m <= 12):
        return False
    if not (1 <= d <= 31):
        return False
    return True


def _pick_nit_from_cells_after_label(row: tuple, label_col: int) -> str:
    """
    Tras la celda 'NIT Tercero', elige el mejor candidato a NIT entre celdas vecinas.
    No usa el primer bloque de dígitos (podría ser una fecha u otro número).
    """
    best = ""
    for k in range(label_col + 1, min(label_col + 8, len(row))):
        n = _norm_nit(row[k])
        if not _is_valid_nit(n):
            continue
        if _looks_like_yyyymmdd_digits(n):
            continue
        if len(n) > len(best):
            best = n
    return best


def _is_valid_nit(nit: str) -> bool:
    """
    Filtra falsos NIT provenientes de celdas vacías/formulas:
    - vacío
    - solo ceros (ej. "0", "00", "000000")
    """
    n = _norm_nit(nit)
    return bool(n) and any(ch != "0" for ch in n)


def _fmt_date(val) -> str:
    if val is None:
        return ""
    if isinstance(val, datetime):
        return val.strftime("%d/%m/%Y")
    s = str(val).strip()
    if len(s) == 8 and s.isdigit():
        return f"{s[6:8]}/{s[4:6]}/{s[0:4]}"
    return s


def _safe_float(val) -> float:
    if val is None or val == "" or val == "-" or val == "#N/A":
        return 0.0
    try:
        if isinstance(val, str):
            s = val.strip()
            if s == "" or s == "-" or s == "#N/A":
                return 0.0
            s = s.replace(" ", "")
            neg = s.startswith("(") and s.endswith(")")
            if neg:
                s = s[1:-1]
            s = s.replace(".", "").replace(",", ".")
            try:
                f = float(s)
                return -f if neg else f
            except ValueError:
                return 0.0
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def _dedupe_fisicos_rows(rows: List[RowFisicoCompra]) -> List[RowFisicoCompra]:
    """
    Quita líneas duplicadas que el pivot suele repetir (mismo vcto e importes):
    si solo una fila del grupo tiene texto en Notas y el resto va vacío, se conserva
    la que tiene notas. Si varias tienen notas, no se fusionan.
    """
    if len(rows) < 2:
        return rows

    def _cents(x: float) -> int:
        return int(round(float(x) * 100.0))

    def _grp(r: RowFisicoCompra):
        return (r.fecha_vcto, _cents(r.debitos), _cents(r.creditos))

    buckets: Dict[Tuple[str, int, int], List[RowFisicoCompra]] = defaultdict(list)
    for r in rows:
        buckets[_grp(r)].append(r)

    out: List[RowFisicoCompra] = []
    for chunk in buckets.values():
        if len(chunk) == 1:
            out.append(chunk[0])
            continue
        nonempty = [r for r in chunk if (r.notas or "").strip()]
        if len(nonempty) == 1:
            out.append(nonempty[0])
        elif len(nonempty) == 0:
            out.append(max(chunk, key=lambda r: (r.num, r.fecha_docto)))
        else:
            out.extend(chunk)

    out.sort(key=lambda r: (r.fecha_docto, r.num))
    return out


# ──────────────────────────────────────────────────────────────
# NIT map desde hoja BD
# ──────────────────────────────────────────────────────────────
def _load_nit_map(wb) -> Tuple[Dict[str, str], Dict[str, dict]]:
    nit_by_name   = {}
    saldos_by_nit = {}
    bd = wb[SHEET_BD]
    for row in bd.iter_rows(min_row=6, values_only=True):
        nit  = row[0]
        name = row[1]
        nit_str = _norm_nit(nit)
        if not (_is_valid_nit(nit_str) and name):
            continue
        name_key = str(name).strip().upper()
        nit_by_name[name_key] = nit_str
        saldos_by_nit[nit_str] = {
            "saldo_inicial": _safe_float(row[2]),
            "entradas":      _safe_float(row[3]),
            "salidas":       _safe_float(row[4]),
            "saldo_final":   _safe_float(row[5]),
        }
    return nit_by_name, saldos_by_nit


def _resolve_nit(razon: str, nit_by_name: Dict[str, str]) -> str:
    key = razon.strip().upper()
    if key in nit_by_name:
        return nit_by_name[key]
    for k, v in nit_by_name.items():
        if k in key or key in k:
            return v
    return "SIN_NIT"


def _nit_to_name(nit_str: str, nit_by_name: Dict[str, str]) -> str:
    n = _norm_nit(nit_str)
    return next((k for k, v in nit_by_name.items() if v == n), n or nit_str)


# ──────────────────────────────────────────────────────────────
# OP VIGENTES
# Fuente de verdad: solo clientes presentes en la hoja OP VIGENTES.
# No se incluyen clientes "extra" desde BD cuando no existen en el pivot.
# ──────────────────────────────────────────────────────────────
def _parse_op_vigentes(
    wb,
    nit_by_name: Dict[str, str],
) -> Tuple[datetime, datetime, List[ClienteOpVigente]]:
    ws = wb[SHEET_OP_VIGENTES]

    desde = hasta = None
    header_row_idx = 16  # Default fallback
    filter_nit = None

    # Búsqueda dinámica de encabezados y metadatos (primeras 30 filas)
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), start=1):
        if len(row) > 4:
            val_d = str(row[3]).strip() if row[3] else ""
            if val_d == "Desde:":
                desde = row[4]
            elif val_d == "Hasta:":
                hasta = row[4]

            # Buscamos filtro de Cliente: Col B="Cliente", Col C=NIT
            val_b = str(row[1]).strip() if row[1] else ""
            if val_b == "Cliente":
                possible_nit = _norm_nit(row[2])
                if _is_valid_nit(possible_nit):
                    filter_nit = possible_nit

        # Buscamos fila de encabezados
        val_a = str(row[0]).strip() if row[0] else ""
        if val_a == "OPERACIONES VIGENTES":
            header_row_idx = i
            break

    start_row = header_row_idx + 1

    # Paso 1: leer filas de detalle desde la hoja
    clientes_dict: Dict[str, ClienteOpVigente] = {}

    # Si hay un NIT de filtro global, inicializamos el cliente para asegurar que exista
    if filter_nit:
        name_bd = _nit_to_name(filter_nit, nit_by_name)
        if name_bd == filter_nit:
            name_bd = f"CLIENTE {filter_nit}"
        clientes_dict[filter_nit] = ClienteOpVigente(nit=filter_nit, razon_social=name_bd)

    for row in ws.iter_rows(min_row=start_row, values_only=True):
        if not row:
            continue
            
        op = row[0]
        if op is None:
            continue
            
        # Chequeo de fin de tabla
        op_str = str(op).strip()
        if op_str.startswith("Total") or op_str == "Nombre Cliente":
            break

        doc_cruce  = row[1]
        fecha_doc  = row[2]
        razon      = row[3]
        fecha_venc = row[4]
        saldo      = row[5]

        razon_str = str(razon).strip() if razon else "DESCONOCIDO"

        # Resolución de NIT
        if filter_nit:
            # Si hay filtro, forzamos este NIT
            nit = filter_nit
            # Actualizamos nombre si encontramos uno mejor en la fila
            if razon_str != "DESCONOCIDO" and clientes_dict[nit].razon_social.startswith("CLIENTE"):
                clientes_dict[nit].razon_social = razon_str
        else:
            nit = _resolve_nit(razon_str, nit_by_name)

        if nit not in clientes_dict:
            clientes_dict[nit] = ClienteOpVigente(nit=nit, razon_social=razon_str)

        clientes_dict[nit].rows.append(RowOpVigente(
            operacion         = op_str,
            doc_cruce         = int(doc_cruce) if doc_cruce else 0,
            fecha_doc         = _fmt_date(fecha_doc),
            razon_social      = razon_str,
            fecha_vencimiento = _fmt_date(fecha_venc),
            saldo_final       = _safe_float(saldo),
        ))

    return desde, hasta, sorted(clientes_dict.values(), key=lambda c: c.razon_social)


# ──────────────────────────────────────────────────────────────
# FISICOS COMPRAS
# Criterio de inclusión (OR):
#   a) Tiene filas en la hoja FISICOS COMPRAS del Excel
#   b) Tiene entradas o salidas != 0 en BD
# Clientes con solo saldo_inicial/saldo_final sin movimiento
# del período NO se incluyen.
# ──────────────────────────────────────────────────────────────
def _parse_fisicos_compras(
    wb,
    nit_by_name: Dict[str, str],
    saldos_by_nit: Dict[str, dict],
    excel_path: str,                    # ← CAMBIO 2: nuevo parámetro
) -> List[ClienteFisicoCompra]:
    ws = wb[SHEET_FISICOS_COMPRAS]

    def _norm_hdr(v) -> str:
        s = str(v).strip().lower()
        if not s:
            return ""
        s = (s.replace("á", "a").replace("é", "e").replace("í", "i")
               .replace("ó", "o").replace("ú", "u").replace("ñ", "n"))
        for ch in (" ", "_", "-", ".", "/", "\\", "\n", "\r", "\t"):
            s = s.replace(ch, "")
        return s

    def _safe_int(v) -> int:
        if v is None or v == "" or v == "-" or v == "#N/A":
            return 0
        try:
            return int(v)
        except Exception:
            s = str(v).strip().replace(".", "")
            return int(s) if s.isdigit() else 0

    header_row_idx = 28
    idx_nombre = idx_fecha_doc = idx_fecha_vcto = idx_docto = idx_num = idx_notas = idx_debitos = idx_creditos = None
    filter_nit = None

    # Dos pasadas: (1) NIT del filtro del pivot puede estar *debajo* del encabezado o más allá de la fila 80.
    # Antes se hacía una sola pasada y al encontrar "Nombre Cliente" se cortaba el bucle, sin leer NIT Tercero posterior.
    scan_last = min(ws.max_row, 250)

    for row in ws.iter_rows(min_row=1, max_row=scan_last, values_only=True):
        if not row:
            continue
        for j, cell in enumerate(row):
            if str(cell).strip() == "NIT Tercero":
                picked = _pick_nit_from_cells_after_label(row, j)
                if _is_valid_nit(picked):
                    filter_nit = picked
                break

    best_score = -1
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=scan_last, values_only=True), start=1):
        if not row:
            continue
        normed = [_norm_hdr(c) for c in row]
        if "nombrecliente" not in normed:
            continue
        score = sum(
            1 for key in (
                "fechadocto", "fechavcto", "docto", "notas", "debitos", "creditos",
            )
            if key in normed
        )
        if "num" in normed:
            score += 1
        if score >= best_score:
            best_score = score
            header_row_idx = i
            idx_nombre = normed.index("nombrecliente")
            idx_fecha_doc = idx_fecha_vcto = idx_docto = idx_num = idx_notas = idx_debitos = idx_creditos = None
            if "fechadocto" in normed:
                idx_fecha_doc = normed.index("fechadocto")
            if "fechavcto" in normed:
                idx_fecha_vcto = normed.index("fechavcto")
            if "docto" in normed:
                idx_docto = normed.index("docto")
            if "num" in normed:
                idx_num = normed.index("num")
            if "notas" in normed:
                idx_notas = normed.index("notas")
            if "debitos" in normed:
                idx_debitos = normed.index("debitos")
            if "creditos" in normed:
                idx_creditos = normed.index("creditos")

    if idx_nombre is None:
        idx_nombre, idx_fecha_doc, idx_fecha_vcto, idx_docto, idx_num, idx_notas, idx_debitos, idx_creditos = range(8)
    else:
        if idx_fecha_doc is None:
            idx_fecha_doc = idx_nombre + 1
        if idx_fecha_vcto is None:
            idx_fecha_vcto = idx_nombre + 2
        if idx_docto is None:
            idx_docto = idx_nombre + 3
        if idx_num is None:
            idx_num = idx_nombre + 4
        if idx_notas is None:
            idx_notas = idx_nombre + 5
        if idx_debitos is None:
            idx_debitos = idx_nombre + 6
        if idx_creditos is None:
            idx_creditos = idx_nombre + 7

    start_row = header_row_idx + 1

    # Paso 1: recolectar movimientos desde la hoja
    movements: Dict[str, List[RowFisicoCompra]] = {}
    observed_names: Dict[str, str] = {}
    current_nit = None

    for row in ws.iter_rows(min_row=start_row, values_only=True):
        if not row:
            continue

        nombre = row[idx_nombre] if idx_nombre < len(row) else None
        fecha_doc = row[idx_fecha_doc] if idx_fecha_doc < len(row) else None
        fecha_vcto = row[idx_fecha_vcto] if idx_fecha_vcto < len(row) else None
        docto = row[idx_docto] if idx_docto < len(row) else None
        num = row[idx_num] if idx_num < len(row) else None
        notas = row[idx_notas] if idx_notas < len(row) else None
        debitos = row[idx_debitos] if idx_debitos < len(row) else None
        creditos = row[idx_creditos] if idx_creditos < len(row) else None

        nombre_str = str(nombre).strip() if nombre is not None else ""
        if nombre_str == "":
            nombre = None

        if nombre is None:
            if current_nit and fecha_doc is not None:
                movements[current_nit].append(RowFisicoCompra(
                    fecha_docto = _fmt_date(fecha_doc),
                    fecha_vcto  = _fmt_date(fecha_vcto),
                    docto       = str(docto) if docto else "",
                    num         = _safe_int(num),
                    notas       = str(notas) if notas else "",
                    debitos     = _safe_float(debitos),
                    creditos    = _safe_float(creditos),
                ))
            continue

        if nombre_str.lower() == "total general":
            break
        if nombre_str.startswith("Total ") or nombre_str == "Nombre Cliente":
            current_nit = None
            continue

        # Primera fila de un cliente en la hoja
        nit = filter_nit if filter_nit else _resolve_nit(nombre_str, nit_by_name)
        current_nit = nit
        
        if nit not in observed_names or len(nombre_str) > len(observed_names[nit]):
            observed_names[nit] = nombre_str

        if nit not in movements:
            movements[nit] = []
        if fecha_doc is not None:
            movements[nit].append(RowFisicoCompra(
                fecha_docto = _fmt_date(fecha_doc),
                fecha_vcto  = _fmt_date(fecha_vcto),
                docto       = str(docto) if docto else "",
                num         = _safe_int(num),
                notas       = str(notas) if notas else "",
                debitos     = _safe_float(debitos),
                creditos    = _safe_float(creditos),
            ))

    # Paso 2: universo de NITs a incluir
    #   = los que tienen filas en la hoja  (movements dict)
    #   + los del BD con entradas o salidas != 0
    nits_a_incluir: set = set(movements.keys())
    for nit_str, saldos in saldos_by_nit.items():
        if saldos["entradas"] != 0.0 or saldos["salidas"] != 0.0:
            nits_a_incluir.add(nit_str)

    clientes: List[ClienteFisicoCompra] = []
    for nit_str in nits_a_incluir:
        saldos = saldos_by_nit.get(nit_str, {
            "saldo_inicial": 0.0, "entradas": 0.0,
            "salidas": 0.0, "saldo_final": 0.0,
        })
        
        bd_name = _nit_to_name(nit_str, nit_by_name)
        final_name = bd_name
        
        if bd_name == nit_str and nit_str in observed_names:
            final_name = observed_names[nit_str]
        elif nit_str in observed_names:
            obs = observed_names[nit_str]
            if len(obs) > len(bd_name) and not obs.replace(".", "").isdigit():
                final_name = obs

        clientes.append(ClienteFisicoCompra(
            nit           = nit_str,
            razon_social  = final_name,
            saldo_inicial = saldos["saldo_inicial"],
            entradas      = saldos["entradas"],
            salidas       = saldos["salidas"],
            saldo_final   = saldos["saldo_final"],
            rows          = movements.get(nit_str, []),
        ))

    # ── CAMBIO 3: Pivot cache fallback ────────────────────────
    # Clientes que tienen saldos en BD (entradas/salidas != 0) pero
    # no tienen filas de detalle en la hoja visible del pivot.
    # Buscamos sus registros en el XML interno del pivot cache.
    nits_sin_detalle = {c.nit for c in clientes if len(c.rows) == 0}
    if nits_sin_detalle:
        try:
            pivot_records = extract_fisicos_from_pivot_cache(excel_path)
            for c in clientes:
                if c.nit in nits_sin_detalle and c.nit in pivot_records:
                    for pr in pivot_records[c.nit]:
                        c.rows.append(RowFisicoCompra(
                            fecha_docto = _fmt_date(pr.fecha_docto),
                            fecha_vcto  = _fmt_date(pr.fecha_vcto),
                            docto       = pr.tipo_docto,
                            num         = int(pr.num_docto) if pr.num_docto.isdigit() else 0,
                            notas       = pr.notas,
                            debitos     = pr.debitos,
                            creditos    = pr.creditos,
                        ))
        except Exception:
            pass  # Si falla la lectura del pivot cache, seguir sin detalle
    # ── Fin pivot cache fallback ──────────────────────────────

    for c in clientes:
        c.rows = _dedupe_fisicos_rows(c.rows)

    return sorted(clientes, key=lambda c: c.razon_social)


# ──────────────────────────────────────────────────────────────
# Entry point público
# ──────────────────────────────────────────────────────────────
def parse_informe(excel_path: str) -> InformeData:
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    nit_by_name, saldos_by_nit = _load_nit_map(wb)
    desde, hasta, clientes_op = _parse_op_vigentes(wb, nit_by_name)
    clientes_fc = _parse_fisicos_compras(wb, nit_by_name, saldos_by_nit, excel_path)  # ← CAMBIO 3
    return InformeData(
        excel_path               = excel_path,
        desde                    = desde,
        hasta                    = hasta,
        clientes_op_vigentes     = clientes_op,
        clientes_fisicos_compras = clientes_fc,
    )


def parse_cartera(excel_path: str) -> CarteraData:
    """
    Estructura real del Excel CARTERA_OMAS:
      - Hoja de metadatos: nombre contiene "OMAS" → fila 7 col 5 tiene "CORTE: DD/MM/YYYY"
      - Hoja de datos: la otra hoja (ej. "CARTERA FEBRERO")
          Fila 1: encabezados
          Fila 2+: datos con columnas:
            Cliente | Nombre Cliente | Cuenta | Documento |
            Fecha Contabilización | Fecha Vencimiento | DIAS |
            Corriente | De 1 a 30 | De 31 a 60 | De 61 a 90 | De 91 a 9999 | Total
      - La cuenta se lee de col3 de cualquier fila de datos (siempre "16161010")
      - Fallback: extraer cuenta del nombre del archivo (_16161010.xlsx)
    """
    import os, re as _re

    wb = openpyxl.load_workbook(excel_path, data_only=True)

    # ── Helpers internos ───────────────────────────────────────
    def _norm_hdr(v) -> str:
        s = str(v).strip().lower()
        if not s:
            return ""
        s = (s.replace("á", "a").replace("é", "e").replace("í", "i")
               .replace("ó", "o").replace("ú", "u").replace("ñ", "n"))
        for ch in (" ", "_", "-", ".", "/", "\\", "\n", "\r", "\t", ":"):
            s = s.replace(ch, "")
        return s

    def _safe_int(v) -> int:
        if v is None or v == "" or v == "-" or v == "#N/A":
            return 0
        try:
            return int(v)
        except Exception:
            s = str(v).strip()
            s = "".join(ch for ch in s if ch.isdigit())
            return int(s) if s else 0

    def _parse_date_any(v):
        if v is None:
            return None
        if isinstance(v, datetime):
            return v
        s = str(v).strip()
        if not s:
            return None
        m = _re.search(r"(\d{1,2}/\d{1,2}/\d{4})", s)
        if m:
            try:
                return datetime.strptime(m.group(1), "%d/%m/%Y")
            except Exception:
                return None
        m2 = _re.search(r"(\d{4}-\d{2}-\d{2})", s)
        if m2:
            try:
                return datetime.strptime(m2.group(1), "%Y-%m-%d")
            except Exception:
                return None
        return None

    def _as_text(v) -> str:
        if v is None:
            return ""
        if isinstance(v, datetime):
            return v.strftime("%d/%m/%Y")
        if isinstance(v, float) and v.is_integer():
            return str(int(v))
        return str(v).strip()

    def _extract_digits(text: str, min_len: int = 4) -> str:
        digits = "".join(ch for ch in str(text) if ch.isdigit())
        return digits if len(digits) >= min_len else ""

    # ── Identificar hoja de metadatos y hoja de datos ─────────
    ws_meta = None
    ws_data = None

    for name in wb.sheetnames:
        name_up = name.upper()
        if "OMAS" in name_up or "META" in name_up:
            ws_meta = wb[name]
        else:
            ws_data = wb[name]

    # Si no se pudo distinguir, usar primera como datos
    if ws_data is None and ws_meta is None:
        ws_data = wb[wb.sheetnames[0]]
    elif ws_data is None:
        # Solo hay una hoja o todas tienen "OMAS"
        ws_data = ws_meta
        ws_meta = None

    # ── Extraer CORTE desde hoja de metadatos ─────────────────
    corte = None

    if ws_meta is not None:
        for r_idx in range(1, ws_meta.max_row + 1):
            for c_idx in range(1, ws_meta.max_column + 1):
                val = ws_meta.cell(row=r_idx, column=c_idx).value
                cell_str = _as_text(val)
                if "CORTE" in cell_str.upper():
                    parsed = _parse_date_any(cell_str)
                    if parsed:
                        corte = parsed
                        break
                    # Puede estar en celda adyacente
                    for dc in (1, 2, -1):
                        adj = ws_meta.cell(row=r_idx, column=c_idx + dc).value
                        parsed = _parse_date_any(_as_text(adj))
                        if parsed:
                            corte = parsed
                            break
                    if corte:
                        break
            if corte:
                break

    # Fallback: buscar corte en hoja de datos también
    if corte is None:
        for r_idx in range(1, min(20, ws_data.max_row + 1)):
            for c_idx in range(1, ws_data.max_column + 1):
                val = ws_data.cell(row=r_idx, column=c_idx).value
                cell_str = _as_text(val)
                if "CORTE" in cell_str.upper():
                    parsed = _parse_date_any(cell_str)
                    if parsed:
                        corte = parsed
                        break
            if corte:
                break

    # ── Extraer CUENTA ─────────────────────────────────────────
    # Estrategia 1: del nombre del archivo (patrón _NNNNNNN.xlsx)
    cuenta = ""
    fname = os.path.basename(excel_path)
    m_fname = _re.search(r"[-_](\d{5,10})\.xlsx$", fname, _re.IGNORECASE)
    if m_fname:
        cuenta = m_fname.group(1)

    # ── Detectar encabezados en hoja de datos ─────────────────
    header_row_idx = None
    idx_cliente = idx_nombre = idx_cuenta_col = idx_documento = None
    idx_fcont = idx_fvenc = idx_dias = None
    idx_corriente = idx_1_30 = idx_31_60 = idx_61_90 = idx_91_9999 = idx_total = None

    for r_idx, row in enumerate(
        ws_data.iter_rows(min_row=1, max_row=80, values_only=True), start=1
    ):
        normed = [_norm_hdr(c) for c in row]
        # Buscar fila que tenga al menos "cliente" y "documento"
        if "cliente" in normed and "documento" in normed:
            header_row_idx = r_idx
            idx_cliente   = normed.index("cliente")
            if "nombrecliente" in normed:
                idx_nombre = normed.index("nombrecliente")
            elif "nombrecli" in normed:
                idx_nombre = normed.index("nombrecli")
            else:
                # Buscar cualquier variante con "nombre"
                for ni, n in enumerate(normed):
                    if "nombre" in n:
                        idx_nombre = ni
                        break
            if "cuenta" in normed:
                idx_cuenta_col = normed.index("cuenta")
            idx_documento = normed.index("documento")
            if "fechacontabilizacion" in normed:
                idx_fcont = normed.index("fechacontabilizacion")
            if "fechavencimiento" in normed:
                idx_fvenc = normed.index("fechavencimiento")
            if "dias" in normed:
                idx_dias = normed.index("dias")
            if "corriente" in normed:
                idx_corriente = normed.index("corriente")
            if "de1a30" in normed:
                idx_1_30 = normed.index("de1a30")
            if "de31a60" in normed:
                idx_31_60 = normed.index("de31a60")
            if "de61a90" in normed:
                idx_61_90 = normed.index("de61a90")
            if "de91a9999" in normed:
                idx_91_9999 = normed.index("de91a9999")
            if "total" in normed:
                idx_total = normed.index("total")
            break

    if header_row_idx is None:
        raise ValueError(
            f"No se encontró el encabezado de la tabla de cartera en '{excel_path}'."
        )

    # Fallbacks posicionales
    if idx_nombre    is None: idx_nombre    = idx_cliente + 1
    if idx_fcont     is None: idx_fcont     = idx_documento + 1
    if idx_fvenc     is None: idx_fvenc     = idx_documento + 2
    if idx_dias      is None: idx_dias      = idx_documento + 3
    if idx_corriente is None: idx_corriente = idx_documento + 4
    if idx_1_30      is None: idx_1_30      = idx_documento + 5
    if idx_31_60     is None: idx_31_60     = idx_documento + 6
    if idx_61_90     is None: idx_61_90     = idx_documento + 7
    if idx_91_9999   is None: idx_91_9999   = idx_documento + 8
    if idx_total     is None: idx_total     = idx_documento + 9

    # ── Leer filas de datos ────────────────────────────────────
    clientes: Dict[str, ClienteCartera] = {}
    current_nit  = ""
    current_name = ""

    for row in ws_data.iter_rows(min_row=header_row_idx + 1, values_only=True):
        if not row:
            continue

        # Detectar fila de "Total general"
        first_val = str(row[idx_cliente]).strip() if idx_cliente < len(row) and row[idx_cliente] is not None else ""
        if first_val.lower() == "total general":
            break

        # NIT del cliente (col "Cliente")
        possible_nit  = _norm_nit(row[idx_cliente])  if idx_cliente  < len(row) else ""
        possible_name = str(row[idx_nombre]).strip()  if idx_nombre   < len(row) and row[idx_nombre]  is not None else ""

        if possible_nit:
            current_nit  = possible_nit
        if possible_name:
            current_name = possible_name

        # Extraer cuenta de la columna "Cuenta" de la fila (Estrategia 2)
        if not cuenta and idx_cuenta_col is not None and idx_cuenta_col < len(row):
            cuenta_val = _as_text(row[idx_cuenta_col]).strip()
            digits = _extract_digits(cuenta_val, min_len=4)
            if digits:
                cuenta = digits

        if not current_nit:
            continue

        # Registrar cliente
        if current_nit not in clientes:
            clientes[current_nit] = ClienteCartera(
                nit=current_nit,
                razon_social=current_name or current_nit,
            )
        elif current_name and clientes[current_nit].razon_social == current_nit:
            clientes[current_nit].razon_social = current_name

        # Leer campos de la fila
        documento  = _as_text(row[idx_documento])  if idx_documento  < len(row) else ""
        if documento.lower().startswith("total"):
            continue

        fcont      = _fmt_date(row[idx_fcont])      if idx_fcont      < len(row) else ""
        fvenc      = _fmt_date(row[idx_fvenc])      if idx_fvenc      < len(row) else ""
        dias       = _safe_int(row[idx_dias])        if idx_dias       < len(row) else 0
        corriente  = _safe_float(row[idx_corriente]) if idx_corriente  < len(row) else 0.0
        de_1_30    = _safe_float(row[idx_1_30])      if idx_1_30       < len(row) else 0.0
        de_31_60   = _safe_float(row[idx_31_60])     if idx_31_60      < len(row) else 0.0
        de_61_90   = _safe_float(row[idx_61_90])     if idx_61_90      < len(row) else 0.0
        de_91_9999 = _safe_float(row[idx_91_9999])   if idx_91_9999    < len(row) else 0.0
        total      = _safe_float(row[idx_total])     if idx_total      < len(row) else 0.0

        if total == 0.0:
            total = corriente + de_1_30 + de_31_60 + de_61_90 + de_91_9999

        # Ignorar filas completamente vacías
        if not (documento or corriente or de_1_30 or de_31_60 or de_61_90 or de_91_9999 or total):
            continue

        clientes[current_nit].rows.append(RowCartera(
            documento             = documento,
            fecha_contabilizacion = fcont,
            fecha_vencimiento     = fvenc,
            dias                  = dias,
            corriente             = corriente,
            de_1_30               = de_1_30,
            de_31_60              = de_31_60,
            de_61_90              = de_61_90,
            de_91_9999            = de_91_9999,
            total                 = total,
        ))

    return CarteraData(
        excel_path = excel_path,
        corte      = corte,
        cuenta     = cuenta,
        clientes   = sorted(clientes.values(), key=lambda c: c.razon_social),
    )