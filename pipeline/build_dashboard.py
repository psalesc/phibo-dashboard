#!/usr/bin/env python3
"""
Gera o dashboard.html a partir do snapshot da planilha mestre (master.xlsx) e,
opcionalmente, da planilha "Controle Vendas" (fonte adicional de vendas
diárias por vendedor + taxa de conversão).

Uso:
    python3 build_dashboard.py <master.xlsx> [saida.html] [controle_vendas.xlsx] [aba_do_mes] [--from-vendors]

Fonte do mês corrente (ver "Virada de mês" no CLAUDE.md):
- Padrão: abas "Resumo Vendedores"/"Vendas Detalhadas" do master.
- Se o master ainda estiver em outro mês (o Dashboard_AoVivo.gs não virou o
  CONFIG), o mês é montado direto das planilhas de vendedor
  (build_data_from_vendors), desde que samuel.xlsx/marina.xlsx/bruno.xlsx
  estejam na mesma pasta do master. Sem elas, o build ABORTA (exit 2) em vez
  de publicar dados de um mês com o rótulo de outro.
- --from-vendors força esse modo mesmo com o master em dia.

Espera as abas: "Resumo Vendedores" e "Vendas Detalhadas" (produzidas pelo
Dashboard_AoVivo.gs). O template estático (dashboard_template.html) precisa
estar no mesmo diretório deste script.

Se `controle_vendas.xlsx` for informado, lê a aba `aba_do_mes` (padrão:
mês atual em português, ex. "Agosto") no formato usado pelo Pedro desde
nov/2025: blocos "Vendas Diárias - Total/Samuel/Bruno/Marina/Maria" nas
colunas L-AD (linhas 5-35, uma por dia do mês) e um bloco de
leads/vendas/taxa de conversão por vendedor nas linhas 34-37 (colunas A-D).
Esse layout é fixo há vários meses nessa planilha, mas pode mudar — em caso
de dúvida, o script avisa no stdout em vez de falhar silenciosamente.
"""
import sys
import os
import json
import calendar
import datetime
import openpyxl

# Número de dias do mês em processamento (calculado dinamicamente em main() a
# partir do mês/ano reais — NÃO mais fixo em 31, pra funcionar automaticamente
# em qualquer mês, inclusive fevereiro). Mantido aqui só como valor de reserva
# para chamadas diretas de build_controle_vendas() sem passar num_days.
NUM_DAYS_FALLBACK = 31

MESES_PT = ['', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
            'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']

# Colunas fixas do bloco "Vendas Diárias" na planilha Controle Vendas:
# nome -> (coluna da data, coluna do valor diário, coluna do acumulado)
CV_BLOCKS = {
    'total': ('L', 'M', 'N'),
    'Samuel': ('P', 'Q', 'R'),
    'Bruno': ('T', 'U', 'V'),
    'Marina': ('X', 'Y', 'Z'),
    'Maria': ('AB', 'AC', 'AD'),
}
CV_FIRST_DATA_ROW = 5
CV_LAST_DATA_ROW = 35  # cobre até 31 dias


def is_num(v):
    return isinstance(v, (int, float))


def aggregate_daily(records, categoria, vendors, num_days):
    """Agrega uma lista de registros (mesmo shape usado em 'records') por dia
    do mês, pra uma categoria ('Mensal' ou 'Anual/Pacote'). Usado tanto pelo
    mês corrente (a partir de "Vendas Detalhadas") quanto pelos meses
    históricos (a partir da extração direta das planilhas dos vendedores)."""
    totals = [0] * num_days
    qtd_totals = [0] * num_days
    by_vendor = {v: [0] * num_days for v in vendors}
    qtd_by_vendor = {v: [0] * num_days for v in vendors}
    for rec in records:
        if rec['categoria'] != categoria or not rec['data']:
            continue
        parts = str(rec['data']).split('/')
        if len(parts) != 3:
            continue
        try:
            dia = int(parts[0])
        except ValueError:
            continue
        if dia < 1 or dia > num_days:
            continue
        valor = rec['valor'] if is_num(rec['valor']) else 0
        totals[dia - 1] += valor
        qtd_totals[dia - 1] += 1
        if rec['vendedor'] in by_vendor:
            by_vendor[rec['vendedor']][dia - 1] += valor
            qtd_by_vendor[rec['vendedor']][dia - 1] += 1
    return {'totals': totals, 'qtdTotals': qtd_totals, 'byVendor': by_vendor, 'qtdByVendor': qtd_by_vendor}


# ============================================================
# EXTRAÇÃO HISTÓRICA — porta direta de extractVendorData() em
# Dashboard_AoVivo.gs, pra reconstruir um mês fechado (não mais presente em
# "Vendas Detalhadas", que o Apps Script sobrescreve a cada mês) direto das
# planilhas individuais dos vendedores (Samuel/Marina/Bruno). Usado só pelo
# script de captura histórica (build_historical.py) — build_data() acima
# continua sendo a via normal pro mês corrente via master.xlsx.
# ============================================================

def fmt_date(val):
    if isinstance(val, (datetime.datetime, datetime.date)):
        return val.strftime('%d/%m/%Y')
    if isinstance(val, str) and val.strip():
        return val.strip()
    return ''


def find_mensal_tab(wb, mm, yyyy):
    mm_yyyy = mm + yyyy
    for name in wb.sheetnames:
        name_upper = name.strip().upper()
        if (name_upper.startswith('CONVERS') and 'BACKUP' not in name_upper and
                (mm_yyyy in name_upper or f'{mm} {yyyy}' in name_upper or f'{mm}/{yyyy}' in name_upper)):
            return wb[name]
    return None


def find_pacotes_tab(wb, mes_nome_pt, yyyy):
    mes_upper = mes_nome_pt.upper()
    for name in wb.sheetnames:
        name_upper = name.upper()
        if 'PACOTE' in name_upper and mes_upper in name_upper and yyyy in name_upper and 'BACKUP' not in name_upper:
            return wb[name]
    return None


def extract_vendor_month(vendor_path, vendor_name, mm, yyyy, mes_nome_pt):
    """Lê a planilha bruta de um vendedor (Samuel/Marina/Bruno) e extrai os
    registros + KPIs de um mês específico, no mesmo formato que
    extractVendorData() grava no master via Apps Script. Levanta ValueError
    se a aba mensal esperada não existir (mês não fechado ainda ou nome de
    aba fora do padrão)."""
    wb = openpyxl.load_workbook(vendor_path, data_only=True)
    records = []

    ws_mensal = find_mensal_tab(wb, mm, yyyy)
    if ws_mensal is None:
        raise ValueError(f'Aba mensal não encontrada para {vendor_name} ({mm}{yyyy}). '
                          f'Abas disponíveis: {wb.sheetnames}')

    for r in range(2, ws_mensal.max_row + 1):
        nome = ws_mensal.cell(row=r, column=2).value  # B
        if nome is None or str(nome).strip() == '':
            continue
        valor = ws_mensal.cell(row=r, column=7).value  # G
        records.append({
            'vendedor': vendor_name,
            'data': fmt_date(ws_mensal.cell(row=r, column=8).value),  # H
            'cliente': nome,
            'categoria': 'Mensal',
            'plano': '',
            'valor': valor if is_num(valor) else None,
            'status': ws_mensal.cell(row=r, column=5).value,  # E
            'telefone': ws_mensal.cell(row=r, column=3).value,  # C
            'observacao': ws_mensal.cell(row=r, column=6).value,  # F
        })

    # bloco de resumo já calculado pela própria planilha (colunas J/K, linhas 1-8)
    summary = {}
    for rr in range(1, 9):
        label = ws_mensal.cell(row=rr, column=10).value  # J
        val = ws_mensal.cell(row=rr, column=11).value  # K
        if label:
            summary[str(label).strip()] = val

    fechamentos = ws_mensal.cell(row=4, column=15).value  # O4

    # linha "Vendedor" (fallback pro Faturamento/Realizado quando a label
    # "Faturamento" não existe no bloco de resumo, como acontece com o Bruno)
    realizado_fallback = None
    for rr2 in range(1, 35):
        if ws_mensal.cell(row=rr2, column=10).value == 'Vendedor':
            realizado_fallback = ws_mensal.cell(row=rr2 + 1, column=13).value  # M
            break

    faturamento = summary.get('Faturamento') if is_num(summary.get('Faturamento')) else realizado_fallback
    meta = summary.get('Meta') if is_num(summary.get('Meta')) else None

    kpi = {
        'mensaisR': summary['Total de vendas'] if 'Total de vendas' in summary else summary.get('Mensais'),
        'anualPacotesR': summary['Total pacotes'] if 'Total pacotes' in summary else summary.get('Pacotes'),
        'faturamentoR': faturamento,
        'metaR': meta,
        'comissaoR': summary.get('Comissão'),
        'fechamentosQtd': fechamentos if is_num(fechamentos) else 0,
    }

    # aba de pacotes anuais — nome varia (ex: "PACOTES - AGOSTO 2026" na
    # Marina/Samuel vs "PACOTES AGOSTO - 2026" no Bruno)
    ws_pacotes = find_pacotes_tab(wb, mes_nome_pt, yyyy)
    if ws_pacotes is not None:
        max_row = ws_pacotes.max_row
        # tabela principal: A=Data, B=Vendedor, C=Cliente, D=WhatsApp, E=CNPJ,
        # F=ID Sistema, G=Valor, H=Serviço, I=Pagamento(status), J=Link, K=Obs
        for p in range(3, max_row + 1):
            cliente = ws_pacotes.cell(row=p, column=3).value  # C
            if cliente is None or str(cliente).strip() == '':
                continue
            valor = ws_pacotes.cell(row=p, column=7).value  # G
            records.append({
                'vendedor': vendor_name,
                'data': fmt_date(ws_pacotes.cell(row=p, column=1).value),  # A
                'cliente': cliente,
                'categoria': 'Anual/Pacote',
                'plano': ws_pacotes.cell(row=p, column=8).value,  # H
                'valor': valor if is_num(valor) else None,
                'status': ws_pacotes.cell(row=p, column=9).value,  # I
                'telefone': ws_pacotes.cell(row=p, column=4).value,  # D
                'observacao': ws_pacotes.cell(row=p, column=11).value,  # K
            })
        # sub-tabela de migração: P-Y (sempre curta, até a linha 29)
        for m in range(3, min(max_row, 29) + 1):
            cliente_m = ws_pacotes.cell(row=m, column=17).value  # Q
            if cliente_m is None or str(cliente_m).strip() == '':
                continue
            valor = ws_pacotes.cell(row=m, column=21).value  # U
            plano = ws_pacotes.cell(row=m, column=22).value or 'Migração'  # V
            obs = ws_pacotes.cell(row=m, column=25).value or 'Migração'  # Y
            records.append({
                'vendedor': vendor_name,
                'data': fmt_date(ws_pacotes.cell(row=m, column=16).value),  # P
                'cliente': cliente_m,
                'categoria': 'Anual/Pacote',
                'plano': plano,
                'valor': valor if is_num(valor) else None,
                'status': ws_pacotes.cell(row=m, column=23).value,  # W
                'telefone': ws_pacotes.cell(row=m, column=18).value,  # R
                'observacao': obs,
            })

    return records, kpi


def build_data_from_vendors(vendor_paths, mm, yyyy, mes_nome_pt, num_days, updated_at_label):
    """Reconstrói o mesmo shape de build_data() direto das planilhas
    individuais dos vendedores, pra um mês fechado que já não está mais
    presente em "Vendas Detalhadas" do master (sobrescrita a cada virada de
    mês). vendor_paths: {'Samuel': 'samuel.xlsx', 'Marina': ..., 'Bruno': ...}
    (mesma ordem em que devem aparecer no filtro de vendedor)."""
    vendors = list(vendor_paths.keys())
    all_records = []
    kpis = {}
    for vendor_name, path in vendor_paths.items():
        recs, kpi = extract_vendor_month(path, vendor_name, mm, yyyy, mes_nome_pt)
        all_records.extend(recs)
        kpis[vendor_name] = kpi

    def sum_field(field):
        return sum((kpis[v].get(field) or 0) for v in vendors)

    metas = [kpis[v].get('metaR') for v in vendors]
    meta_total = sum(m for m in metas if is_num(m)) if all(is_num(m) for m in metas) else None
    faturamento_total = sum_field('faturamentoR')
    fechamentos_total = sum_field('fechamentosQtd')
    total = {
        'mensaisR': sum_field('mensaisR'),
        'anualPacotesR': sum_field('anualPacotesR'),
        'faturamentoR': faturamento_total,
        'fechamentosQtd': fechamentos_total,
        'ticketMedio': (faturamento_total / fechamentos_total) if fechamentos_total else 0,
        'metaR': meta_total,
        'pctMeta': (faturamento_total / meta_total) if meta_total else None,
        'comissaoR': sum_field('comissaoR'),
    }
    for v in vendors:
        k = kpis[v]
        fech = k.get('fechamentosQtd') or 0
        fat = k.get('faturamentoR') or 0
        k['ticketMedio'] = (fat / fech) if fech else 0
        m = k.get('metaR')
        k['pctMeta'] = (fat / m) if is_num(m) and m else None

    daily = {
        'mensal': aggregate_daily(all_records, 'Mensal', vendors, num_days),
        'pacotes': aggregate_daily(all_records, 'Anual/Pacote', vendors, num_days),
    }

    return {
        'updatedAt': updated_at_label,
        'vendors': vendors,
        'kpis': kpis,
        'total': total,
        'daily': daily,
        'records': all_records,
    }


def recompute_fechamentos_from_records(data):
    """Corrige fechamentosQtd (e o ticketMedio que depende dele) contando as
    linhas reais de 'records' por vendedor, em vez de confiar no campo
    'Fechamentos'/'Total Mensal' que cada vendedor calcula na própria
    planilha (Resumo Vendedores, no mês corrente; O4 de cada aba mensal, no
    histórico).

    Achado (Pedro, 15/09): o KPI "Fechamentos" do painel não batia com o
    número de linhas na aba Transações do mesmo vendedor. Causa: esse campo
    vem de uma fórmula de intervalo fixo dentro da planilha de cada
    vendedor (ex. "Total Mensal"), que não acompanha vendas lançadas depois
    que o intervalo foi definido — historicamente ficou ~5 registros atrás
    pro Samuel no fechamento de agosto, e no dia 15/09 estava 3 registros
    atrás pro Samuel (parado em 88, com 91 linhas reais na aba do mês) e 2
    atrás pra Marina, além de estar 1 registro À FRENTE pro Bruno (a
    fórmula soma uma linha que não existe mais na aba). Contar as linhas
    reais de 'records' elimina essa dependência de qualquer fórmula interna
    de vendedor e garante que o KPI sempre bate com a tabela de Transações
    e com os exports."""
    if not data.get('records'):
        return
    from collections import Counter
    counts = Counter(r['vendedor'] for r in data['records'])
    for v in data.get('vendors', []):
        if v not in data['kpis']:
            continue
        fech = counts.get(v, 0)
        data['kpis'][v]['fechamentosQtd'] = fech
        fat = data['kpis'][v].get('faturamentoR') or 0
        data['kpis'][v]['ticketMedio'] = (fat / fech) if fech else 0
    total_fech = len(data['records'])
    data['total']['fechamentosQtd'] = total_fech
    total_fat = data['total'].get('faturamentoR') or 0
    data['total']['ticketMedio'] = (total_fat / total_fech) if total_fech else 0


def reconcile_kpis_with_cv(data, cv):
    """Corrige faturamentoR/metaR/pctMeta/ticketMedio usando a planilha
    Controle Vendas como fonte de verdade, quando disponível.

    Achado (Pedro, 14/09): o "Faturamento Total" de agosto no painel batia
    478k, mas o valor real (Controle Vendas) é 537k — uma diferença de ~59k.
    Investigando: o campo "Faturamento" que cada vendedor reporta na própria
    planilha (usado por extract_vendor_month/build_data para o mês corrente
    também, via "Resumo Vendedores") não é confiável pra todo mundo — bate
    exatamente pra Marina, mas fica ~37k abaixo pro Samuel e ~22k abaixo pro
    Bruno (que nem tem o rótulo "Faturamento" na própria planilha, dependendo
    de um fallback). A Controle Vendas — preenchida manualmente pelo Pedro —
    é o registro que reflete a realidade; cruzando com o mês corrente
    (setembro/2026) confirma-se que ela bate em cheio com "Resumo
    Vendedores" quando os dois existem lado a lado. Por isso, sempre que a
    Controle Vendas estiver disponível (mês corrente ou histórico), ela
    prevalece sobre o valor "Faturamento" que vem da planilha do vendedor.

    Achado (Pedro, 25/09): a Meta total mostrada no painel batia R$480.000,
    mas o valor correto é R$482.000. Causa: a Meta de cada vendedor também
    vem da planilha do próprio vendedor (via "Resumo Vendedores"), que pode
    ficar desatualizada quando a meta muda no meio do mês — a planilha do
    Samuel ainda tinha a meta antiga de R$175.000, enquanto a Controle
    Vendas já refletia os R$177.000 corretos (bloco-resumo "Vendedor /
    Conversões Mensais / Anual-Pacotes / Realizado / Média diária / Meta /
    Faltam / % / Vendas por dia" no topo de cada aba mensal). Mesma lógica
    do Faturamento acima: a Controle Vendas prevalece sobre a Meta que a
    planilha do vendedor guarda, sempre que disponível.

    Não mexe em mensaisR/anualPacotesR/comissaoR (a Controle Vendas não tem
    essa quebra) nem nos registros usados pela tabela de transações/gráfico
    por plano — só nos números de faturamento/meta usados nos KPIs."""
    if not cv or not cv.get('dailyByVendor'):
        return

    # Meta: a Controle Vendas prevalece sobre a Meta que a planilha de cada
    # vendedor guarda (achado Pedro, 25/09 — ver comentário acima e em
    # build_controle_vendas). Aplicado ANTES de recalcular pctMeta, senão o
    # percentual usaria a meta antiga.
    metas_cv = cv.get('metasByVendor') or {}
    for v, meta_val in metas_cv.items():
        if v in data['kpis'] and is_num(meta_val):
            data['kpis'][v]['metaR'] = meta_val
    meta_total_cv = cv.get('metaTotalCV')
    if is_num(meta_total_cv):
        data['total']['metaR'] = meta_total_cv

    # Faturamento POR VENDEDOR: coluna "Realizado" da tabela-resumo da
    # Controle Vendas (a que o Pedro confere e o painel antigo usa). A soma
    # dos blocos "Vendas Diárias - <vendedor>" fica só como reserva — achado
    # Pedro, 15/10/2026: em outubro esses blocos diários por vendedor estavam
    # ~R$22k atrás do Realizado (Bruno 40.219 vs 48.691; Samuel 59.876 vs
    # 69.756), enquanto o bloco "Total" diário batia. O TOTAL continua vindo
    # de dailyTotal (valores de ago/set já validados).
    realizado_cv = cv.get('realizadoByVendor') or {}
    for v, series in cv['dailyByVendor'].items():
        if v not in data['kpis']:
            continue
        fat = realizado_cv.get(v)
        if not is_num(fat):
            fat = round(sum(series), 2)
        fat = round(fat, 2)
        data['kpis'][v]['faturamentoR'] = fat
        meta = data['kpis'][v].get('metaR')
        data['kpis'][v]['pctMeta'] = (fat / meta) if is_num(meta) and meta else None
        fech = data['kpis'][v].get('fechamentosQtd') or 0
        data['kpis'][v]['ticketMedio'] = (fat / fech) if fech else 0
    # TOTAL: linha "Total" da coluna Realizado da tabela-resumo, quando tem
    # valor (achado 15/10/2026: em outubro a tabela-resumo era atualizada ao
    # longo do dia e o bloco diário "Total" ficava para trás). Em ago/set essa
    # célula está vazia, então cai no bloco diário — que é de onde vêm os
    # valores já validados (R$537.617,54 / R$533.168,43).
    total_fat = cv.get('realizadoTotalCV')
    if not is_num(total_fat):
        total_fat = sum(cv['dailyTotal'])
    total_fat = round(total_fat, 2)
    data['total']['faturamentoR'] = total_fat
    meta_total = data['total'].get('metaR')
    data['total']['pctMeta'] = (total_fat / meta_total) if is_num(meta_total) and meta_total else None
    fech_total = data['total'].get('fechamentosQtd') or 0
    data['total']['ticketMedio'] = (total_fat / fech_total) if fech_total else 0


def build_data(xlsx_path, num_days):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    ws = wb['Vendas Detalhadas']
    # "Vendas Detalhadas" ganhou 5 colunas novas na nona rodada (Telefone,
    # CNPJ/CPF, ID Sistema, Link de Pagamento, Comprovante) — detecta pelo
    # cabeçalho da coluna 8 em vez de assumir, pra continuar funcionando com
    # um snapshot antigo (8 colunas, Observação na coluna 8) se algum dia
    # for preciso reprocessar um.
    header_col8 = str(ws.cell(row=1, column=8).value or '').strip().lower()
    has_telefone = header_col8.startswith('telefone')
    records = []
    for r in range(2, ws.max_row + 1):
        vendedor = ws.cell(row=r, column=1).value
        if vendedor is None or str(vendedor).strip() == '':
            continue
        records.append({
            'vendedor': vendedor,
            'data': ws.cell(row=r, column=2).value,
            'cliente': ws.cell(row=r, column=3).value,
            'categoria': ws.cell(row=r, column=4).value,
            'plano': ws.cell(row=r, column=5).value,
            'valor': ws.cell(row=r, column=6).value,
            'status': ws.cell(row=r, column=7).value,
            'telefone': ws.cell(row=r, column=8).value if has_telefone else None,
            'observacao': ws.cell(row=r, column=13).value if has_telefone else ws.cell(row=r, column=8).value,
        })

    wsr = wb['Resumo Vendedores']
    vendors = []
    kpis = {}
    r = 6
    while True:
        v = wsr.cell(row=r, column=2).value
        if not v or v == 'Time (Total)':
            break
        vendors.append(v)
        kpis[v] = {
            'mensaisR': wsr.cell(row=r, column=3).value or 0,
            'anualPacotesR': wsr.cell(row=r, column=4).value or 0,
            'faturamentoR': wsr.cell(row=r, column=5).value or 0,
            'fechamentosQtd': wsr.cell(row=r, column=6).value or 0,
            'ticketMedio': wsr.cell(row=r, column=7).value or 0,
            'metaR': wsr.cell(row=r, column=8).value,
            'pctMeta': wsr.cell(row=r, column=9).value if is_num(wsr.cell(row=r, column=9).value) else None,
            'comissaoR': wsr.cell(row=r, column=10).value or 0,
        }
        r += 1
    total_row = r
    total = {
        'mensaisR': wsr.cell(row=total_row, column=3).value or 0,
        'anualPacotesR': wsr.cell(row=total_row, column=4).value or 0,
        'faturamentoR': wsr.cell(row=total_row, column=5).value or 0,
        'fechamentosQtd': wsr.cell(row=total_row, column=6).value or 0,
        'ticketMedio': wsr.cell(row=total_row, column=7).value or 0,
        'metaR': wsr.cell(row=total_row, column=8).value,
        'pctMeta': wsr.cell(row=total_row, column=9).value if is_num(wsr.cell(row=total_row, column=9).value) else None,
        'comissaoR': wsr.cell(row=total_row, column=10).value or 0,
    }
    updated_at = wsr.cell(row=3, column=2).value

    daily = {
        'mensal': aggregate_daily(records, 'Mensal', vendors, num_days),
        'pacotes': aggregate_daily(records, 'Anual/Pacote', vendors, num_days),
    }

    return {
        'updatedAt': updated_at,
        'vendors': vendors,
        'kpis': kpis,
        'total': total,
        'daily': daily,
        'records': records,
    }


def build_controle_vendas(xlsx_path, month_tab=None, meta_total=None, num_days=None):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)

    today = datetime.date.today()
    if month_tab is None:
        month_tab = MESES_PT[today.month]

    if month_tab not in wb.sheetnames:
        print(f'AVISO: aba "{month_tab}" não encontrada em Controle Vendas (abas disponíveis: {wb.sheetnames}). Pulando essa fonte.')
        return None
    ws = wb[month_tab]

    # dias reais do mês (28-31 conforme mês/ano) — calculado dinamicamente pra
    # funcionar automaticamente em qualquer mês, sem precisar editar o script.
    if num_days is None:
        month_num_for_days = MESES_PT.index(month_tab)
        num_days = calendar.monthrange(today.year, month_num_for_days)[1]
    NUM_DAYS = num_days

    # aviso não-fatal se os títulos dos blocos mudaram de posição/nome
    expected_titles = {
        'L3': 'Vendas Diárias - Total', 'P3': 'Vendas Diárias - Samuel',
        'T3': 'Vendas Diárias - Bruno', 'X3': 'Vendas Diárias - Marina',
        'AB3': 'Vendas Diárias - Maria',
    }
    for cell_ref, expected in expected_titles.items():
        # o título pode estar ancorado 1 coluna à direita (mescla) — checa a vizinhança
        col = openpyxl.utils.column_index_from_string(''.join(ch for ch in cell_ref if ch.isalpha()))
        neighborhood = [ws.cell(row=3, column=c).value for c in range(col, col + 2)]
        if expected not in neighborhood:
            print(f'AVISO: cabeçalho esperado "{expected}" não encontrado perto de {cell_ref} na aba "{month_tab}" — layout pode ter mudado, conferir manualmente.')

    days = list(range(1, NUM_DAYS + 1))
    daily = {}
    dates_seen = None
    for name, (dcol, vcol, _ccol) in CV_BLOCKS.items():
        dcol_i = openpyxl.utils.column_index_from_string(dcol)
        vcol_i = openpyxl.utils.column_index_from_string(vcol)
        series = []
        block_dates = []
        for i, r in enumerate(range(CV_FIRST_DATA_ROW, CV_FIRST_DATA_ROW + NUM_DAYS)):
            d = ws.cell(row=r, column=dcol_i).value
            v = ws.cell(row=r, column=vcol_i).value
            block_dates.append(d.day if isinstance(d, (datetime.datetime, datetime.date)) else (i + 1))
            v = v if is_num(v) else 0
            series.append(v)
        daily[name] = series
        if name == 'total':
            dates_seen = block_dates

    # guarda contra o bug de fórmula da planilha original: no dia da virada
    # ("hoje"), algumas colunas de vendedor voltam um valor negativo igual a
    # -1x o acumulado do dia anterior. Vendas diárias nunca são negativas —
    # qualquer valor negativo é tratado como dado ruim (zerado) e sinalizado.
    anomaly_days = set()
    for name, series in daily.items():
        for i, v in enumerate(series):
            if v < 0:
                anomaly_days.add(i + 1)
                series[i] = 0

    # dias já decorridos no mês (comparando com a data real de hoje) — só usa
    # o dia atual se a aba corresponde ao mês corrente; se for um mês passado,
    # considera o mês inteiro decorrido.
    if MESES_PT[today.month] == month_tab:
        elapsed = today.day
    else:
        elapsed = NUM_DAYS
    elapsed = max(1, min(elapsed, NUM_DAYS))

    total_series = daily['total']
    elapsed_vals = total_series[:elapsed]
    active_vals = [v for v in elapsed_vals if v > 0]
    stats = {
        'melhorDia': max(elapsed_vals) if elapsed_vals else 0,
        'mediaDiaAtivo': (sum(active_vals) / len(active_vals)) if active_vals else 0,
        'diasComVenda': len(active_vals),
        'diasElapsed': elapsed,
    }

    cumulative_real = []
    running = 0
    for i, v in enumerate(total_series):
        if i < elapsed:
            running += v
            cumulative_real.append(round(running, 2))
        else:
            cumulative_real.append(None)

    projecao_meta = None
    if meta_total:
        projecao_meta = [round(meta_total * d / NUM_DAYS, 2) for d in days]

    # tabela-resumo "Vendedor / Conversões Mensais / Anual-Pacotes / Realizado /
    # Média diária / Meta / Faltam / % / Vendas por dia" (linhas ~4-8, colunas
    # A-I) — é aqui que a Meta de cada vendedor fica de fato atualizada; a
    # Meta que cada vendedor guarda na própria planilha (usada em
    # extract_vendor_month/"Resumo Vendedores") pode ficar desatualizada
    # quando a meta muda no meio do mês (achado Pedro, 25/09: total do painel
    # mostrava R$480.000, mas o correto é R$482.000 — a planilha do Samuel
    # ainda tinha a meta antiga de R$175.000, enquanto a Controle Vendas já
    # refletia os R$177.000 corretos). Localizada por busca de texto (não
    # linha fixa) pra resistir a pequenos deslocamentos de layout.
    metas_by_vendor = {}
    realizado_by_vendor = {}
    realizado_total_cv = None
    meta_total_cv = None
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(row=r, column=1).value or '').strip().lower() == 'vendedor':
            meta_col = None
            realizado_col = None
            for c in range(1, 12):
                label = str(ws.cell(row=r, column=c).value or '').strip().lower()
                if label == 'meta' and meta_col is None:
                    meta_col = c
                elif label == 'realizado' and realizado_col is None:
                    realizado_col = c
            if meta_col:
                rr = r + 1
                while True:
                    nome = ws.cell(row=rr, column=1).value
                    if not nome:
                        break
                    meta_val = ws.cell(row=rr, column=meta_col).value
                    if str(nome).strip().lower() == 'total':
                        meta_total_cv = meta_val if is_num(meta_val) else None
                        if realizado_col:
                            rt = ws.cell(row=rr, column=realizado_col).value
                            realizado_total_cv = rt if is_num(rt) else None
                        break
                    metas_by_vendor[str(nome).strip()] = meta_val if is_num(meta_val) else None
                    if realizado_col:
                        real_val = ws.cell(row=rr, column=realizado_col).value
                        if is_num(real_val):
                            realizado_by_vendor[str(nome).strip()] = real_val
                    rr += 1
            break

    # leads / vendas / taxa de conversão (linhas 34-37, colunas A-D)
    conversion = {}
    for r in range(34, 38):
        nome = ws.cell(row=r, column=1).value
        if not nome:
            continue
        leads = ws.cell(row=r, column=2).value
        vendas = ws.cell(row=r, column=3).value
        taxa = ws.cell(row=r, column=4).value
        conversion[nome] = {
            'leads': leads if is_num(leads) else 0,
            'vendas': vendas if is_num(vendas) else 0,
            'taxa': taxa if is_num(taxa) else 0,
        }

    # bloco "RENOVAÇÕES" (Maria) — Realizado/Média diária/Meta/Faltam/% em uma
    # tabelinha separada, geralmente logo abaixo do bloco de vendas diárias.
    # Localizada por busca de texto em vez de linha fixa, porque a posição já
    # variou entre abas de meses diferentes nesta planilha.
    renovacoes = None
    for r in range(1, ws.max_row + 1):
        if str(ws.cell(row=r, column=1).value or '').strip().upper() == 'RENOVAÇÕES':
            header_row = r + 1
            data_row = r + 2
            nome = ws.cell(row=data_row, column=1).value
            if nome and str(nome).strip().lower().startswith('mari'):
                realizado = ws.cell(row=data_row, column=2).value
                media_diaria = ws.cell(row=data_row, column=3).value
                meta = ws.cell(row=data_row, column=4).value
                faltam = ws.cell(row=data_row, column=5).value
                pct = ws.cell(row=data_row, column=6).value
                renovacoes = {
                    'realizadoR': realizado if is_num(realizado) else 0,
                    'mediaDiariaR': media_diaria if is_num(media_diaria) else 0,
                    'metaR': meta if is_num(meta) else 0,
                    'faltamR': faltam if is_num(faltam) else 0,
                    'pctMeta': pct if is_num(pct) else None,
                }
            break
    if renovacoes is None:
        print(f'AVISO: bloco "RENOVAÇÕES" (Maria) não encontrado na aba "{month_tab}" — a aba Renovações do painel vai ficar sem os KPIs de topo.')

    maria_series = daily['Maria']
    maria_elapsed_vals = maria_series[:elapsed]
    maria_active_vals = [v for v in maria_elapsed_vals if v > 0]
    maria_stats = {
        'melhorDia': max(maria_elapsed_vals) if maria_elapsed_vals else 0,
        'mediaDiaAtivo': (sum(maria_active_vals) / len(maria_active_vals)) if maria_active_vals else 0,
        'diasComVenda': len(maria_active_vals),
        'diasElapsed': elapsed,
    }

    # fins de semana do mês (para sombrear o gráfico de evolução acumulada).
    # usa o ano corrente se a aba for o mês corrente, senão assume o ano
    # corrente também (a planilha não guarda o ano por aba) — é só um detalhe
    # visual, não afeta nenhum número.
    month_num = MESES_PT.index(month_tab)
    year = today.year
    weekends = []
    for d in days:
        try:
            wd = datetime.date(year, month_num, d).weekday()  # 0=seg ... 6=dom
            weekends.append(wd >= 5)
        except ValueError:
            weekends.append(False)

    return {
        'monthTab': month_tab,
        'weekends': weekends,
        'days': days,
        'dailyTotal': daily['total'],
        'dailyByVendor': {k: v for k, v in daily.items() if k not in ('total', 'Maria')},
        'dailyMaria': daily['Maria'],
        'cumulativeReal': cumulative_real,
        'projecaoMeta': projecao_meta,
        'stats': stats,
        'conversion': conversion,
        'anomalyDays': sorted(anomaly_days),
        'renovacoes': renovacoes,
        'renovacoesStats': maria_stats,
        'metasByVendor': metas_by_vendor,
        'realizadoByVendor': realizado_by_vendor,
        'realizadoTotalCV': realizado_total_cv,
        'metaTotalCV': meta_total_cv,
    }


def dominant_record_month(records):
    """(mm, yyyy) mais frequente nas datas dd/mm/yyyy dos records, ou None."""
    from collections import Counter
    c = Counter()
    for r in records or []:
        parts = str(r.get('data') or '').split('/')
        if len(parts) == 3 and len(parts[2]) == 4:
            c[(parts[1].zfill(2), parts[2])] += 1
    return c.most_common(1)[0][0] if c else None


def vendor_paths_next_to(master_path):
    """Planilhas de vendedor na mesma pasta do master, ou None se faltar alguma."""
    base = os.path.dirname(os.path.abspath(master_path))
    paths = {v: os.path.join(base, f'{v.lower()}.xlsx') for v in ('Samuel', 'Marina', 'Bruno')}
    return paths if all(os.path.exists(p) for p in paths.values()) else None


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)
    args = [a for a in sys.argv[1:] if a != '--from-vendors']
    force_vendors = '--from-vendors' in sys.argv[1:]
    xlsx_path = args[0]
    out_path = args[1] if len(args) > 1 else 'dashboard.html'
    controle_vendas_path = args[2] if len(args) > 2 else None
    month_tab = args[3] if len(args) > 3 else None
    script_dir = os.path.dirname(os.path.abspath(__file__))
    template_path = os.path.join(script_dir, 'dashboard_template.html')

    # Mês/ano em processamento: por padrão o mês corrente real (assim o painel
    # "rola" sozinho de mês em mês sem precisar editar nada aqui); o 4º
    # argumento (aba_do_mes) ainda permite forçar um mês específico, por
    # exemplo pra reprocessar um mês fechado. Dias do mês calculados de
    # verdade (28-31) em vez de fixos em 31, pra fevereiro e meses de 30 dias
    # ficarem corretos também.
    today = datetime.date.today()
    if month_tab:
        month_num = MESES_PT.index(month_tab)
    else:
        month_num = today.month
        month_tab = MESES_PT[month_num]
    year = today.year
    num_days = calendar.monthrange(year, month_num)[1]
    mes_ano = f'{month_tab} {year}'
    mes_num_padded = f'{month_num:02d}'

    data = None
    if not force_vendors:
        data = build_data(xlsx_path, num_days)
        master_month = dominant_record_month(data.get('records'))
        if master_month and master_month != (mes_num_padded, str(year)):
            print(f'AVISO: o master está em {master_month[0]}/{master_month[1]}, mas o mês '
                  f'processado é {mes_num_padded}/{year} — o Dashboard_AoVivo.gs ainda não '
                  f'virou o CONFIG. Montando o mês a partir das planilhas de vendedor.')
            data = None
    if data is None:
        vendor_paths = vendor_paths_next_to(xlsx_path)
        if not vendor_paths:
            print('ERRO: master desatualizado e planilhas de vendedor (samuel/marina/bruno.xlsx) '
                  'não encontradas ao lado dele. Abortando para não publicar dados do mês errado.')
            sys.exit(2)
        updated_label = 'Atualizado automaticamente em ' + datetime.datetime.now().strftime('%d/%m/%Y, %H:%M:%S')
        data = build_data_from_vendors(vendor_paths, mes_num_padded, str(year), month_tab, num_days, updated_label)
        print(f'INFO: mês {mes_num_padded}/{year} montado das planilhas de vendedor ({len(data["records"])} registros).')
    data['mesAno'] = mes_ano
    data['mesNumPadded'] = mes_num_padded
    recompute_fechamentos_from_records(data)

    if controle_vendas_path:
        meta_total = data['total'].get('metaR') or 0
        cv = build_controle_vendas(controle_vendas_path, month_tab, meta_total, num_days)
        data['controleVendas'] = cv
        reconcile_kpis_with_cv(data, cv)
    else:
        data['controleVendas'] = None

    # Meses históricos: cache JSON gerado uma vez por build_historical.py
    # (extração direta das planilhas dos vendedores, já que "Vendas
    # Detalhadas" no master só guarda o mês corrente). Cada arquivo em
    # historical/*.json já vem no shape completo de build_data(), com
    # mesAno/monthNum/year adicionados. O mês corrente nunca é sobrescrito
    # por um cache — só os meses fechados aparecem aqui.
    historical_dir = os.path.join(script_dir, 'historical')
    historical_blobs = []
    if os.path.isdir(historical_dir):
        for fname in sorted(os.listdir(historical_dir)):
            if not fname.endswith('.json'):
                continue
            fpath = os.path.join(historical_dir, fname)
            try:
                with open(fpath, encoding='utf-8') as f:
                    blob = json.load(f)
                if blob.get('mesAno') == mes_ano:
                    continue  # nunca deixa um cache velho sobrepor o mês corrente
                historical_blobs.append(blob)
            except (json.JSONDecodeError, OSError) as e:
                print(f'AVISO: não foi possível ler cache histórico "{fname}": {e}')
    historical_blobs.sort(key=lambda b: (b.get('year', 0), b.get('monthNum', 0)), reverse=True)

    months_order = [mes_ano] + [b['mesAno'] for b in historical_blobs]
    months = {mes_ano: data}
    for b in historical_blobs:
        months[b['mesAno']] = b
    months_bundle = {'order': months_order, 'months': months}

    with open(template_path, encoding='utf-8') as f:
        template = f.read()

    marker = '/*__DATA_JSON__*/'
    if marker not in template:
        print('ERRO: marcador de dados não encontrado no template.')
        sys.exit(1)
    html = template.replace(marker, json.dumps(months_bundle, ensure_ascii=False))

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

    cv_note = 'sem Controle Vendas' if not data['controleVendas'] else f'Controle Vendas aba={data["controleVendas"]["monthTab"]}'
    hist_note = f', históricos={[b["mesAno"] for b in historical_blobs]}' if historical_blobs else ''
    print(f'OK: {out_path} gerado ({len(html)} bytes). vendors={data["vendors"]} atualizado_em={data["updatedAt"]} ({cv_note}){hist_note}')


if __name__ == '__main__':
    main()
