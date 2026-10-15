#!/usr/bin/env python3
"""
Captura única de um mês fechado direto das planilhas dos vendedores (não
depende do master.xlsx, que só guarda o mês corrente em "Vendas
Detalhadas"). Gera historical/<mes>_<ano>.json, que build_dashboard.py
carrega automaticamente em toda build seguinte.

Uso:
    python3 build_historical.py <mes_pt> <yyyy> <mm> \
        <samuel.xlsx> <marina.xlsx> <bruno.xlsx> \
        <controle_vendas.xlsx> [updated_at_label]

Exemplo (Agosto/2026):
    python3 build_historical.py Agosto 2026 08 \
        samuel.xlsx marina.xlsx bruno.xlsx controle_vendas.xlsx \
        "Agosto/2026 (fechado)"
"""
import sys
import os
import json
import calendar

from build_dashboard import (
    build_data_from_vendors, build_controle_vendas, reconcile_kpis_with_cv,
    recompute_fechamentos_from_records, MESES_PT,
)


def main():
    if len(sys.argv) < 8:
        print(__doc__)
        sys.exit(1)
    mes_nome_pt = sys.argv[1]
    yyyy = sys.argv[2]
    mm = sys.argv[3]
    samuel_path = sys.argv[4]
    marina_path = sys.argv[5]
    bruno_path = sys.argv[6]
    controle_vendas_path = sys.argv[7]
    updated_at_label = sys.argv[8] if len(sys.argv) > 8 else f'{mes_nome_pt}/{yyyy} (fechado)'

    month_num = MESES_PT.index(mes_nome_pt)
    year = int(yyyy)
    num_days = calendar.monthrange(year, month_num)[1]

    vendor_paths = {'Samuel': samuel_path, 'Marina': marina_path, 'Bruno': bruno_path}
    data = build_data_from_vendors(vendor_paths, mm, yyyy, mes_nome_pt, num_days, updated_at_label)
    recompute_fechamentos_from_records(data)

    meta_total = data['total'].get('metaR') or 0
    cv = build_controle_vendas(controle_vendas_path, mes_nome_pt, meta_total, num_days)
    data['controleVendas'] = cv
    reconcile_kpis_with_cv(data, cv)

    data['mesAno'] = f'{mes_nome_pt} {year}'
    data['mesNumPadded'] = f'{month_num:02d}'
    data['monthNum'] = month_num
    data['year'] = year

    script_dir = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(script_dir, 'historical')
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f'{mes_nome_pt.lower()}_{year}.json')
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False)

    cv_note = 'sem Controle Vendas' if not cv else f'Controle Vendas aba={cv["monthTab"]}'
    print(f'OK: {out_path} gerado. vendors={data["vendors"]} total.faturamentoR={data["total"]["faturamentoR"]} ({cv_note})')


if __name__ == '__main__':
    main()
