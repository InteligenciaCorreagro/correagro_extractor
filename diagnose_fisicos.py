"""
Diagnóstico específico para NIT 890301163 (DISTRIBUIDORA COLOMBINA LTDA).
Busca en todos los pivot caches resolviendo índices correctamente.
Ejecutar: python diagnose_890301163.py "INFORME EXTRACTOS MARZO 2026.xlsx"
"""
import sys
import zipfile
import xml.etree.ElementTree as ET

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

def _find(p, n):
    r = p.find(f'{NS}{n}')
    return r if r is not None else p.find(n)

def _find_all(p, n):
    r = p.findall(f'{NS}{n}')
    return r if r else p.findall(n)

def run(excel_path, target_nit="890301163"):
    with zipfile.ZipFile(excel_path, 'r') as z:
        all_files = z.namelist()
        cache_defs = sorted([f for f in all_files if 'pivotCacheDefinition' in f and f.endswith('.xml') and '_rels' not in f])
        cache_recs = sorted([f for f in all_files if 'pivotCacheRecords' in f and f.endswith('.xml')])

        for def_file, rec_file in zip(cache_defs, cache_recs):
            print(f"\n{'='*100}")
            print(f"Cache: {def_file} + {rec_file}")

            cd_xml = z.read(def_file).decode('utf-8')
            cd_root = ET.fromstring(cd_xml)
            cf_parent = _find(cd_root, 'cacheFields')
            if cf_parent is None:
                print("  No cacheFields"); continue
            cache_fields = _find_all(cf_parent, 'cacheField')

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

            # Find Cliente field
            cliente_idx = None
            for i, fn in enumerate(field_names):
                if fn.lower().strip() == 'cliente':
                    cliente_idx = i; break
            if cliente_idx is None:
                print("  No 'Cliente' field"); continue

            # Find target NIT index in shared items
            target_si = None
            for si_idx, (tag, val) in enumerate(field_shared[cliente_idx]):
                if val == target_nit:
                    target_si = si_idx; break
            if target_si is None:
                print(f"  NIT {target_nit} NOT in shared items of Cliente field")
                # Check if it appears anywhere
                for i, fn in enumerate(field_names):
                    for si_idx, (tag, val) in enumerate(field_shared[i]):
                        if target_nit in str(val) or 'COLOMBINA' in str(val).upper():
                            print(f"    Found in field[{i}] '{fn}': {tag}={val} (index={si_idx})")
                continue

            print(f"  Cliente field[{cliente_idx}], target shared index={target_si}")

            # Read records
            cr_xml = z.read(rec_file).decode('utf-8')
            cr_root = ET.fromstring(cr_xml)
            records = _find_all(cr_root, 'r')
            print(f"  Total records: {len(records)}")

            for rec_idx, rec in enumerate(records):
                children = list(rec)
                if len(children) <= cliente_idx:
                    continue
                ch = children[cliente_idx]
                ch_tag = ch.tag.replace(NS, '')
                ch_val = ch.get('v', '')

                if ch_tag == 'x' and ch_val == str(target_si):
                    # Decode full record
                    print(f"\n  *** Record #{rec_idx} ***")
                    for fi, child in enumerate(children):
                        if fi >= len(field_names):
                            break
                        c_tag = child.tag.replace(NS, '')
                        c_val = child.get('v', child.text or '')
                        fname = field_names[fi]

                        if c_tag == 'x':
                            try:
                                idx = int(c_val)
                                if idx < len(field_shared[fi]):
                                    _, resolved = field_shared[fi][idx]
                                    # Only show non-empty relevant fields
                                    if resolved and fname.lower() in (
                                        'nombre_auxiliar', 'auxiliar', 'cliente', 'nombre cliente',
                                        'fecha_docto', 'fecha_documento', 'fecha docto fecha',
                                        'fecha_vcto', 'fecha_vencimiento', 'fecha_docto_cruce',
                                        'tipo_docto_cruce', 'tipo documento',
                                        'numero_docto', 'numero_docto_cruce',
                                        'notas', 'debitos', 'creditos',
                                        'periodo_cargos', 'periodo_abonos',
                                        'cuenta', 'nombre_cuenta_n5',
                                    ):
                                        print(f"    [{fi}] {fname}: {resolved}")
                                elif fname.lower() in ('nombre_auxiliar', 'auxiliar', 'nombre cliente'):
                                    print(f"    [{fi}] {fname}: x={c_val} (index)")
                            except ValueError:
                                pass
                        elif c_tag == 'n' and float(c_val) != 0:
                            if fname.lower() in (
                                'debitos', 'creditos', 'periodo_cargos', 'periodo_abonos',
                                'numero_docto', 'numero_docto_cruce',
                                'saldo_inicial', 'saldo_final',
                            ):
                                print(f"    [{fi}] {fname}: {c_val}")
                        elif c_tag == 's' and c_val:
                            if fname.lower() in (
                                'fecha_docto', 'fecha_vcto', 'fecha_vencimiento',
                                'fecha_docto_cruce', 'notas',
                            ):
                                print(f"    [{fi}] {fname}: {c_val}")

if __name__ == "__main__":
    run(sys.argv[1] if len(sys.argv) > 1 else "INFORME EXTRACTOS MARZO 2026.xlsx")