"""
Lector del Pivot Cache interno del Excel (.xlsx).

Cuando el pivot table de la hoja "FISICOS COMPRAS" filtra ciertos clientes,
sus filas no aparecen en las celdas visibles (openpyxl no las ve).
Sin embargo, los datos crudos permanecen en el XML del pivot cache
dentro del archivo ZIP del .xlsx.

Este módulo extrae esos registros ocultos para complementar el parser principal.
"""
import zipfile
import xml.etree.ElementTree as ET
from typing import List, Dict, Tuple
from dataclasses import dataclass

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'


@dataclass
class PivotFisicoRecord:
    """Registro individual de FISICOS COMPRAS extraído del pivot cache."""
    nit: str
    nombre_cliente: str
    fecha_docto: str
    fecha_vcto: str
    tipo_docto: str
    num_docto: str
    notas: str
    debitos: float
    creditos: float
    auxiliar: str


def _find(parent, local_name):
    r = parent.find(f'{NS}{local_name}')
    return r if r is not None else parent.find(local_name)


def _find_all(parent, local_name):
    results = parent.findall(f'{NS}{local_name}')
    return results if results else parent.findall(local_name)


def _safe_float(val: str) -> float:
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def _parse_cache_pair(
    z: zipfile.ZipFile,
    def_file: str,
    rec_file: str,
) -> Tuple[List[str], List[List[Tuple[str, str]]], List[dict]]:
    cd_xml = z.read(def_file).decode('utf-8')
    cd_root = ET.fromstring(cd_xml)

    cf_parent = _find(cd_root, 'cacheFields')
    if cf_parent is None:
        return [], [], []
    cache_fields = _find_all(cf_parent, 'cacheField')
    if not cache_fields:
        return [], [], []

    field_names = []
    field_shared = []

    for cf in cache_fields:
        field_names.append(cf.get('name', ''))
        shared = []
        si = _find(cf, 'sharedItems')
        if si is not None:
            for child in si:
                tag = child.tag.replace(NS, '')
                val = child.get('v', child.text or '')
                shared.append((tag, val))
        field_shared.append(shared)

    cr_xml = z.read(rec_file).decode('utf-8')
    cr_root = ET.fromstring(cr_xml)
    records = _find_all(cr_root, 'r')

    decoded_records = []
    for rec in records:
        children = list(rec)
        decoded = {}
        for fi, ch in enumerate(children):
            if fi >= len(field_names):
                break
            ch_tag = ch.tag.replace(NS, '')
            ch_val = ch.get('v', ch.text or '')

            if ch_tag == 'x':
                try:
                    idx = int(ch_val)
                    if idx < len(field_shared[fi]):
                        _, resolved = field_shared[fi][idx]
                        decoded[field_names[fi]] = resolved
                    else:
                        decoded[field_names[fi]] = ''
                except (ValueError, IndexError):
                    decoded[field_names[fi]] = ''
            elif ch_tag in ('n', 's', 'd'):
                decoded[field_names[fi]] = ch_val
            elif ch_tag == 'm':
                decoded[field_names[fi]] = ''
            else:
                decoded[field_names[fi]] = ch_val

        decoded_records.append(decoded)

    return field_names, field_shared, decoded_records


def extract_fisicos_from_pivot_cache(
    excel_path: str,
) -> Dict[str, List[PivotFisicoRecord]]:
    """
    Lee TODOS los registros de FISICOS COMPRAS del pivot cache del Excel.
    Retorna dict: NIT → lista de PivotFisicoRecord.
    """
    result: Dict[str, List[PivotFisicoRecord]] = {}
    seen: set = set()  # Dedup: (nit, num_docto, fecha_docto, debitos, creditos)

    try:
        with zipfile.ZipFile(excel_path, 'r') as z:
            all_files = z.namelist()
            cache_defs = sorted([
                f for f in all_files
                if 'pivotCacheDefinition' in f
                and f.endswith('.xml')
                and '_rels' not in f
            ])
            cache_recs = sorted([
                f for f in all_files
                if 'pivotCacheRecords' in f
                and f.endswith('.xml')
            ])

            # Procesar en orden inverso: los caches con más campos de detalle
            # (como Notas) tienden a estar en los archivos con número mayor.
            for def_file, rec_file in zip(reversed(cache_defs), reversed(cache_recs)):
                field_names, field_shared, decoded_records = _parse_cache_pair(
                    z, def_file, rec_file
                )
                if not field_names:
                    continue

                fn_lower_set = {fn.lower().strip() for fn in field_names}
                if 'cliente' not in fn_lower_set:
                    continue

                def _get(record, *candidates, default=''):
                    for c in candidates:
                        for key in record:
                            if key.lower().strip() == c.lower().strip():
                                val = record[key]
                                if val is not None and str(val).strip():
                                    return str(val).strip()
                    return default

                for rec in decoded_records:
                    auxiliar = _get(rec, 'Nombre_auxiliar', 'Auxiliar')
                    auxiliar_upper = auxiliar.upper()
                    # La hoja FISICOS COMPRAS incluye ambas cuentas:
                    #   71653105 = FISICOS COMPRAS
                    #   71653108 = MERCOP VENTAS
                    if 'FISICOS COMPRAS' not in auxiliar_upper and 'MERCOP VENTAS' not in auxiliar_upper:
                        continue

                    nit = _get(rec, 'Cliente')
                    if not nit:
                        continue

                    fecha_docto = _get(
                        rec, 'Fecha_docto', 'Fecha_documento',
                        'Fecha docto fecha',
                    )
                    num_docto = _get(rec, 'Numero_docto', 'Numero_docto_cruce')
                    debitos = _safe_float(
                        _get(rec, 'Debitos', 'Periodo_cargos', default='0')
                    )
                    creditos = _safe_float(
                        _get(rec, 'Creditos', 'Periodo_abonos', default='0')
                    )

                    # Ignorar registros sin movimiento (débitos=0 y créditos=0)
                    if debitos == 0.0 and creditos == 0.0:
                        continue

                    notas = _get(rec, 'Notas')

                    # Deduplicar: mismo NIT + num_docto + fecha + montos = mismo registro
                    dedup_key = (nit, num_docto, fecha_docto, debitos, creditos)
                    if dedup_key in seen:
                        continue
                    seen.add(dedup_key)

                    pr = PivotFisicoRecord(
                        nit=nit,
                        nombre_cliente=_get(rec, 'Nombre Cliente'),
                        fecha_docto=fecha_docto,
                        fecha_vcto=_get(rec, 'Fecha_vcto', 'Fecha_vencimiento'),
                        tipo_docto=_get(rec, 'Tipo_docto_cruce', 'Tipo Documento'),
                        num_docto=num_docto,
                        notas=_get(rec, 'Notas'),
                        debitos=debitos,
                        creditos=creditos,
                        auxiliar=auxiliar,
                    )

                    if nit not in result:
                        result[nit] = []
                    result[nit].append(pr)

    except (zipfile.BadZipFile, FileNotFoundError, KeyError):
        pass

    return result