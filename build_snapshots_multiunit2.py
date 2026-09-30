import os
import time, urllib.request, urllib.error, base64, json, datetime, sys, io, calendar, re
import openpyxl

API_BASE_ROOT = "https://evo-integracao.w12app.com.br"

_REQUEST_TIMES = []
_RATE_LIMIT_PER_SEC = 4
_RATE_LIMIT_PER_MIN = 35
_MAX_RETRIES = 6


def _rate_limit_wait():
    now = time.time()
    while _REQUEST_TIMES and now - _REQUEST_TIMES[0] > 60:
        _REQUEST_TIMES.pop(0)
    if len(_REQUEST_TIMES) >= _RATE_LIMIT_PER_MIN:
        wait = 60 - (now - _REQUEST_TIMES[0]) + 0.1
        if wait > 0:
            time.sleep(wait)
        now = time.time()
        while _REQUEST_TIMES and now - _REQUEST_TIMES[0] > 60:
            _REQUEST_TIMES.pop(0)
    recent = [t for t in _REQUEST_TIMES if now - t < 1]
    if len(recent) >= _RATE_LIMIT_PER_SEC:
        wait = 1 - (now - recent[0]) + 0.1
        if wait > 0:
            time.sleep(wait)
    _REQUEST_TIMES.append(time.time())

UNITS = [
    {"key": "fitness", "label": "Fitness", "dns": "4liveacademia", "token": os.environ["EVO_TOKEN_FITNESS"], "suffix": ""},
    {"key": "piscina", "label": "Piscina", "dns": "4liveacademia", "token": os.environ["EVO_TOKEN_PISCINA"], "suffix": "Piscina"},
]


def make_api_get_raw(dns, token):
    auth = base64.b64encode(f"{dns}:{token}".encode()).decode()

    def api_get_raw(url):
        last_err = None
        for attempt in range(_MAX_RETRIES):
            _rate_limit_wait()
            req = urllib.request.Request(url)
            req.add_header("Authorization", f"Basic {auth}")
            try:
                with urllib.request.urlopen(req, timeout=60) as resp:
                    return resp.read()
            except urllib.error.HTTPError as e:
                last_err = e
                if e.code == 429 or 500 <= e.code < 600:
                    retry_after = e.headers.get("Retry-After") if e.headers else None
                    try:
                        wait = float(retry_after) if retry_after else (2 ** attempt)
                    except Exception:
                        wait = 2 ** attempt
                    print(f"[rate-limit] HTTP {e.code} em {url[:90]} — tentativa {attempt+1}/{_MAX_RETRIES}, aguardando {wait:.1f}s", file=sys.stderr)
                    time.sleep(wait)
                    continue
                raise
            except (urllib.error.URLError, TimeoutError, ConnectionError) as e:
                last_err = e
                wait = 2 ** attempt
                print(f"[network-retry] {e} em {url[:90]} — tentativa {attempt+1}/{_MAX_RETRIES}, aguardando {wait:.1f}s", file=sys.stderr)
                time.sleep(wait)
                continue
        raise last_err

    return api_get_raw


def get_recent_entries(api_get_raw, days=90, recent_days=7, all_dates_out=None):
    end = datetime.datetime.now()
    start = end - datetime.timedelta(days=days)
    take = 10000
    skip = 0
    last_by_member = {}
    recent_dates_by_member = {}
    cutoff = (end - datetime.timedelta(days=recent_days - 1)).date().isoformat()
    while True:
        url = f"{API_BASE_ROOT}/api/v1/entries?registerDateStart={start.isoformat()}&registerDateEnd={end.isoformat()}&take={take}&skip={skip}"
        data = json.loads(api_get_raw(url))
        for e in data:
            idm = e.get("idMember")
            d = e.get("date")
            if not idm or not d:
                continue
            d10 = d[:10]
            if idm not in last_by_member or d10 > last_by_member[idm]:
                last_by_member[idm] = d10
            if d10 >= cutoff:
                recent_dates_by_member.setdefault(idm, set()).add(d10)
            if all_dates_out is not None:
                all_dates_out.setdefault(idm, set()).add(d10)
        if len(data) < take:
            break
        skip += take
    visits_recent_by_member = {idm: len(dates) for idm, dates in recent_dates_by_member.items()}
    return last_by_member, visits_recent_by_member


def parse_freq_contratada(plano_name):
    if not plano_name:
        return "Outros"
    m = re.search(r'(\d)\s*X\b', plano_name, re.IGNORECASE)
    return f"{m.group(1)}x" if m else "Outros"


def get_member_activity_sessions(api_get_raw, idm, date_start, date_end):
    url = (
        f"{API_BASE_ROOT}/api/v2/activities/member/sessions?idMember={idm}"
        f"&dateStart={date_start.isoformat()}&dateEnd={date_end.isoformat()}&take=200&skip=0"
    )
    return json.loads(api_get_raw(url))


def build_contract_swap_report(api_get_raw, api_get, api_base_mgmt, api_base_v2, today, month_start):
    anchor = month_start
    for _ in range(2):
        anchor = (anchor - datetime.timedelta(days=1)).replace(day=1)
    range_start, range_end = anchor, today

    nr = load_rows(api_get(api_base_mgmt, "/not-renewed", {"dtStart": range_start.isoformat(), "dtEnd": range_end.isoformat()}))
    cancelled_raw = [r for r in nr if str(r.get("FlCancelado")).strip().lower() == "true"]
    seen = {}
    for r in cancelled_raw:
        try:
            idc = int(r.get("IdCliente"))
        except Exception:
            continue
        dtc = parse_date_br(r.get("DtCancelamento"))
        if not dtc:
            continue
        key = (idc, dtc)
        if key not in seen:
            seen[key] = r

    fetch_end = range_end + datetime.timedelta(days=1)
    sales_raw = []
    skip = 0
    while True:
        chunk = json.loads(api_get_raw(f"{api_base_v2}/sales?dateSaleStart={range_start.isoformat()}&dateSaleEnd={fetch_end.isoformat()}&take=1000&skip={skip}"))
        sales_raw.extend(chunk)
        if len(chunk) < 1000:
            break
        skip += 1000

    sales_by_id = {}
    by_member = {}
    for s in sales_raw:
        if s.get("removed"):
            continue
        sd = s.get("saleDate")
        try:
            sd_d = datetime.datetime.fromisoformat(sd).date() if sd else None
        except Exception:
            sd_d = None
        idm = s.get("idMember")
        items = s.get("saleItens", [])
        total = sum((it.get("saleValue") or 0) for it in items)
        planos = [it.get("item") for it in items]
        has_adesao = any("ades" in (it.get("item") or "").lower() for it in items)
        sales_by_id[s["idSale"]] = {"idMember": idm, "date": sd_d, "total": total, "planos": planos, "has_adesao": has_adesao}
        if idm:
            by_member.setdefault(idm, []).append((sd_d, s["idSale"]))

    matches = []
    for (idc, dtc), r in seen.items():
        candidates = by_member.get(idc, [])
        best = None
        for sd, idsale in candidates:
            if sd and 0 <= (sd - dtc).days <= 1 and sales_by_id[idsale]["has_adesao"]:
                if best is None or sales_by_id[idsale]["total"] > sales_by_id[best]["total"]:
                    best = idsale
        if best:
            motivo = r.get("MotivoCancelamento") or ""
            motivo_confirma = "troca de contrato" in motivo.lower() or "troca de plano" in motivo.lower()
            best_sale = sales_by_id[best]
            matches.append({
                "idCliente": idc, "nome": f"{r.get('Nome','')} {r.get('Sobrenome','')}".strip(),
                "dtCancelamento": dtc.isoformat(), "contratoCancelado": r.get("ContratoCancelado"),
                "dataVendaNova": best_sale["date"].isoformat(), "valorVendaNova": round(best_sale["total"], 2),
                "planosVendaNova": best_sale["planos"], "motivoConfirmaTroca": motivo_confirma,
                # Pedido do usuario em 2026-09-29: idSale da venda nova, pra build_sales_report()
                # poder excluir essa venda das contagens de Novo/Renovacao/Retorno (ver
                # troca_sale_ids_mes em build_for_unit / "trocaContrato" em build_sales_report).
                "idSaleNova": best,
            })

    matches.sort(key=lambda m: m["dtCancelamento"])
    total_value = round(sum(m["valorVendaNova"] for m in matches), 2)
    this_month_matches = [m for m in matches if m["dataVendaNova"] >= month_start.isoformat()]
    this_month_value = round(sum(m["valorVendaNova"] for m in this_month_matches), 2)

    return {
        "periodoInicio": range_start.isoformat(), "periodoFim": range_end.isoformat(),
        "count": len(matches), "valorTotal": total_value, "items": matches,
        "countMesCorrente": len(this_month_matches), "valorMesCorrente": this_month_value,
        "updatedAt": datetime.datetime.utcnow().isoformat() + "Z",
    }


def load_rows(xlsx_bytes):
    wb = openpyxl.load_workbook(io.BytesIO(xlsx_bytes))
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    header = rows[0]
    return [dict(zip(header, r)) for r in rows[1:]]


def parse_date_br(s):
    if not s:
        return None
    try:
        return datetime.datetime.strptime(str(s).strip(), "%d/%m/%Y").date()
    except Exception:
        return None


def parse_value_br(s):
    if not s:
        return 0.0
    s = str(s).replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def categorize_reason(reason):
    norm_map = {
        "problema de saúde": "Saúde", "problemas de saúde": "Saúde",
        "inadimplência": "Inadimplência", "falta de tempo/não frequenta": "Falta de tempo",
        "mudou-se": "Mudança de cidade", "problemas financeiros": "Financeiro",
    }
    parts = (reason or "").split(" - ")
    if len(parts) >= 2:
        cat = parts[1].strip().split("\n")[0].strip()
        if len(cat) > 40:
            cat = cat[:40].rsplit(" ", 1)[0] + "…"
    else:
        cat = "Não informado"
    lk = cat.lower()
    if lk in norm_map:
        return norm_map[lk]
    if lk.startswith("insatisfação"):
        return "Insatisfação com atendimento"
    if lk.startswith("observações: troca") or "troca de contrato" in lk:
        return "Troca interna de plano"
    return cat if cat else "Não informado"


def month_bounds(d):
    start = d.replace(day=1)
    last_day = calendar.monthrange(d.year, d.month)[1]
    end = d.replace(day=last_day)
    return start, end


def parse_iso_date(s):
    if not s:
        return None
    try:
        return datetime.datetime.fromisoformat(s).date()
    except Exception:
        try:
            return datetime.date.fromisoformat(str(s)[:10])
        except Exception:
            return None


def is_transfer_record(c):
    return (c.get("nameMembership") or "").strip().lower().startswith("transferido")


# Pedido do usuario em 2026-09-29: "bike" removido desta lista — o unico plano do sistema que
# batia nessa palavra-chave era "Bike Mensal" (14 contratos na unidade Fitness, confirmado via
# /api/v3/membermembership; nenhum na Piscina), um plano recorrente real que deve contar no
# relatorio de contratos vencendo/renovados/nao renovados igual aos demais planos.
_RENEWAL_EXCLUDE_KEYWORDS = ["colaborador", "cortesia", "diaria", "permuta", "personal", "transferencia"]


# Pedido do usuario em 2026-09-29: ajustes MANUAIS de tipo de venda (Novo/Renovacao/Retorno),
# pontuais, por (unidade, idMember, data da venda) -> (tipoVenda, renovacaoAntecipada). Nao e
# regra geral — NAO adicionar casos aqui sem o usuario pedir.
# - Maira Oliveira Souza (Fitness, idMember 5746, venda 12524 em 25/09/2026, Fitness Assinatura
#   Plus): o unico contrato anterior no Fitness e um registro "Transferido da filial 4LIVE PISCINA"
#   (fim 20/09/2026), que a regra ignora, entao caia como Novo. Usuario pediu Renovacao.
# Pedido do usuario em 2026-09-29: vendas de setembro/2026 da Piscina que o usuario confirmou
# serem TROCA DE CONTRATO (nao devem contar em Novo/Renovacao/Retorno), por idSale:
# Leticia Bonfim Correia de Oliveira (12395), Hugo Fonseca Rosseto (12396), Leonardo Leite Bastos
# (12398, 12406, 12531), Marcelo Leite Bastos (12399, 12405), Yasmin Silva Oliveira (12400, 12402).
# Pedido do usuario em 2026-09-30: mais 3 trocas de contrato na Piscina em setembro/2026 —
# Joao Vicente Ribeiro Menezes (12470), Lunna Maria Melo Farias (12421), Miguel Pacheco Oliveira
# (12388, venda nova com Adesao em 01/09 13:24; a 12375 de 01/09 00:00 e a cobranca recorrente e
# continua como renovacao).
_MANUAL_TROCA_SALE_IDS = {
    "piscina": {12395, 12396, 12398, 12399, 12400, 12402, 12405, 12406, 12531, 12388, 12421, 12470},
}

_MANUAL_SALE_TYPE_OVERRIDES = {
    ("fitness", 5746, datetime.date(2026, 9, 25)): ("renovacao", False),
}


def _strip_accents(s):
    import unicodedata
    return "".join(ch for ch in unicodedata.normalize("NFD", s) if unicodedata.category(ch) != "Mn")


def is_excluded_from_renewals(plano_name):
    nl = _strip_accents((plano_name or "").strip().lower())
    return any(k in nl for k in _RENEWAL_EXCLUDE_KEYWORDS)


# Pedido do usuario em 2026-09-29: contratos que "vencem" mas na verdade foram uma TROCA DE
# CONTRATO/MODALIDADE (o cliente trocou de plano, nao decidiu renovar ou nao) nao devem aparecer
# em "Contratos vencendo" — o vencimento real passa a ser o do contrato novo (que aparece
# naturalmente no mes em que ELE vence, sem nenhum tratamento especial). Regra validada
# nome-a-nome pelo usuario com 3 casos reais na Piscina em setembro/2026: Joao Vicente Ribeiro
# Menezes ("Infantil 1x Anual" -> "Infantil 2x Anual", 1 dia de intervalo), Lunna Maria Melo
# Farias ("Infantil 2x Assinatura" -> "Infantil Sexta Assinatura", 1 dia), Leonardo Leite Bastos
# ("Infantil Sexta Assinatura" -> "Infantil 2x Assinatura", 1 dia). Criterio: existe outro contrato
# do mesmo cliente, NAO cancelado, comecando no mesmo dia ou no dia seguinte ao fim deste (0-1 dia
# de gap — mesmo criterio ja usado em build_contract_swap_report para casar cancelamento + nova
# venda), com um nome de plano DIFERENTE (normalizado, sem acento/maiusculas) — indicando troca de
# modalidade, nao uma renovacao/decisao real do cliente. Uma continuacao do MESMO plano (ciclo
# normal de assinatura, ex. Lunna mes a mes ate a troca) NAO bate nessa regra e continua aparecendo
# normalmente. NAO expandir esse criterio (ex. gaps maiores, contratos cancelados contando como
# "troca") sem validar com o usuario — ver nota equivalente em build_contract_swap_report.
def is_internal_plan_swap(c, contracts):
    plano_atual = _strip_accents((c.get("nameMembership") or "").strip().lower())
    end = c["_end"]
    if end is None:
        return False
    for x in contracts:
        if x is c or is_transfer_record(x) or x.get("statusMemberMembership") != 1:
            continue
        if not x["_start"]:
            continue
        gap = (x["_start"] - end).days
        if 0 <= gap <= 1 and _strip_accents((x.get("nameMembership") or "").strip().lower()) != plano_atual:
            return True
    return False


# Pedido do usuario em 2026-09-29 (mesma correcao acima): Leonardo Leite Bastos (idMember 3576,
# Piscina) tinha DOIS vencimentos dentro de setembro/2026 — um em 25/09 (troca de plano, ja pego
# por is_internal_plan_swap acima) e outro em 05/09, mesmo nome de plano ("Infantil Sexta
# Assinatura") no contrato seguinte, entao NAO bate no criterio de "plano diferente". Nos dados
# brutos do EVO ha 3 registros de contrato sobrepostos/duplicados em torno de 05-06/09 pra esse
# cliente (um deles criado E cancelado no mesmo dia, 06/09) — sinal de correcao administrativa,
# nao um ciclo normal de renovacao (diferente do padrao limpo de Lunna, por exemplo, sem
# sobreposicao nem contrato cancelado no meio). O usuario confirmou que esse tambem nao e uma
# renovacao real e deve sumir de "Contratos vencendo". Como o padrao exato (sobreposicao +
# contrato duplicado cancelado no mesmo dia) so foi validado pra este caso, a exclusao aqui e
# pontual (por idMember + inicio do contrato) em vez de um criterio geral novo — generalizar isso
# exigiria validar mais casos com o usuario (mesma cautela do paragrafo acima). Se aparecerem
# casos parecidos no futuro, checar com o usuario antes de expandir pra uma regra geral.
_MANUAL_RENEWAL_EXCLUSIONS = {
    (3576, "2026-08-06"),  # Leonardo Leite Bastos — Piscina — "Infantil Sexta Assinatura" Ago-Set
}


def is_manually_excluded_from_renewals(idm, c):
    start = c.get("_start")
    if start is None:
        return False
    return (idm, start.isoformat()) in _MANUAL_RENEWAL_EXCLUSIONS


def categorize_sale_item(name):
    n = (name or "").strip()
    nl = n.lower()
    if "fitness" in nl:
        return "Fitness (Anual/Assinatura/Plus)"
    if "adesão" in nl or "adesao" in nl:
        return "Adesão"
    if "colaborador" in nl:
        return "Colaborador"
    if "bike" in nl:
        return "Bike"
    if "assinatura" in nl or "anual" in nl or "mensal" in nl:
        return "Piscina (recorrente)"
    return n if n else "Outros"


def build_loyalty_report(category_filter, contracts_by_member, TODAY, gap_fiel_days=30, gap_retorno_days=60):
    fieis = []
    retornaram = []
    parcela_items = []

    for idm, contracts in contracts_by_member.items():
        core = [c for c in contracts if not is_transfer_record(c) and categorize_sale_item(c.get("nameMembership")) == category_filter]
        if not core:
            continue
        core.sort(key=lambda c: c["_start"] or datetime.date.min)

        active_now = [
            c for c in core
            if c.get("statusMemberMembership") == 1 and c["_start"] and c["_start"] <= TODAY
            and (c["_end"] is None or TODAY <= c["_end"])
        ]
        if not active_now:
            continue
        current = max(active_now, key=lambda c: c["_start"])

        running_max_end = None
        max_gap = 0
        for c in core:
            if running_max_end is None:
                running_max_end = c["_endEfetivo"] or datetime.date.max
                continue
            gap = (c["_start"] - running_max_end).days if c["_start"] else 0
            if gap > max_gap:
                max_gap = gap
            running_max_end = max(running_max_end, c["_endEfetivo"] or datetime.date.max)

        first_start = core[0]["_start"]
        tenure_days = (TODAY - first_start).days if first_start else 0
        nome = current.get("name")
        valor_atual = round(current.get("saleValue") or 0, 2)
        base_item = {
            "idCliente": idm, "nome": nome, "desde": first_start.isoformat() if first_start else None,
            "tenureDias": tenure_days, "planoAtual": (current.get("nameMembership") or "").strip(),
            "valorAtual": valor_atual,
        }

        if max_gap <= gap_fiel_days:
            fieis.append(base_item)
        if max_gap > gap_retorno_days:
            retornaram.append(dict(base_item, maiorIntervaloDias=max_gap))

        parcela = round(valor_atual / 12, 2)
        parcela_items.append({"idCliente": idm, "nome": nome, "parcela": parcela, "planoAtual": base_item["planoAtual"]})

    fieis.sort(key=lambda x: x["tenureDias"], reverse=True)
    retornaram.sort(key=lambda x: x["maiorIntervaloDias"], reverse=True)

    buckets = {}
    for it in parcela_items:
        b = int(it["parcela"] // 50) * 50
        buckets.setdefault(b, []).append(it)
    bucket_list = [
        {"min": b, "max": b + 50, "label": f"R$ {b}-{b+50}", "count": len(items), "items": items}
        for b, items in sorted(buckets.items())
    ]
    total_parcela_clients = len(parcela_items)
    avg_parcela = round(sum(x["parcela"] for x in parcela_items) / total_parcela_clients, 2) if total_parcela_clients else 0

    return {
        "categoria": category_filter, "gapFielDias": gap_fiel_days, "gapRetornoDias": gap_retorno_days,
        "fieis": {"count": len(fieis), "items": fieis},
        "retornaram": {"count": len(retornaram), "items": retornaram},
        "parcelas": {"totalClients": total_parcela_clients, "avgParcela": avg_parcela, "buckets": bucket_list},
        "updatedAt": datetime.datetime.utcnow().isoformat() + "Z",
    }


# Pedido do usuario em 2026-09-28: no dashboard, clicar nos tiles de composicao da base (Ativos,
# VIP, Suspensos, Ex-clientes, Oportunidades, Total, Cancelamentos, Leads) e ver a lista de NOMES.
# Gravado num documento separado ("clientLists"/"clientListsPiscina", doc_id "atual", sobrescrito
# todo dia) para nao inchar o snapshot diario. Cancelados e leads do mes tambem vao para o
# "clientReports" do mes (congelam no fechamento, como os outros campos mensais).
def _fmt_iso_br(d):
    return d.strftime("%d/%m/%Y") if d else None


def build_client_lists(TODAY, active, all_members_full, contracts_by_member, current_contract_for,
                       vip_ids, suspended_ids, all_time_prospects, cancelled, prospects):
    def member_name(m):
        return " ".join(x for x in [m.get("firstName"), m.get("lastName")] if x).strip() or None
    names = {m["idMember"]: member_name(m) for m in all_members_full}

    def contract_info(idm):
        nome_c, fim = current_contract_for(idm)
        return nome_c or None, (_fmt_iso_br(datetime.date.fromisoformat(fim)) if fim else None)

    ativos = []
    for r in active:
        try:
            idc = int(r.get("IdCliente"))
        except Exception:
            idc = None
        ativos.append({"id": idc, "nome": (r.get("NomeCompleto") or names.get(idc) or "").strip() or None,
                       "contrato": (r.get("ContratoAtivo") or "").strip() or None,
                       "data": r.get("DtFimContratoAtivo") or None})

    def from_ids(ids):
        out = []
        for idm in ids:
            c, fim = contract_info(idm)
            out.append({"id": idm, "nome": names.get(idm), "contrato": c, "data": fim})
        return out

    ex = []
    for m in all_members_full:
        if m.get("status") != "Inactive" or m["idMember"] not in contracts_by_member:
            continue
        cs = [c for c in contracts_by_member[m["idMember"]] if not is_transfer_record(c)] or contracts_by_member[m["idMember"]]
        last = max(cs, key=lambda c: c.get("_endEfetivo") or c.get("_start") or datetime.date.min)
        ex.append({"id": m["idMember"], "nome": member_name(m),
                   "contrato": (last.get("nameMembership") or "").strip() or None,
                   "data": _fmt_iso_br(last.get("_endEfetivo"))})

    oportunidades = [{"id": r.get("IdProspect"), "nome": (r.get("Nome") or "").strip() or None, "data": r.get("DtCadastro") or None}
                     for r in all_time_prospects if str(r.get("Status")).strip().upper() == "PROSPECT"]

    cancelados = [{"id": r.get("IdCliente"), "nome": f"{r.get('Nome') or ''} {r.get('Sobrenome') or ''}".strip() or None,
                   "contrato": r.get("ContratoCancelado") or r.get("ContratoVencido") or None,
                   "data": r.get("DtCancelamento") or None, "motivo": categorize_reason(r.get("MotivoCancelamento")),
                   "valor": round(parse_value_br(r.get("Valor")), 2)} for r in cancelled]

    leads = [{"id": r.get("IdProspect"), "nome": (r.get("Nome") or "").strip() or None,
              "status": "Convertido" if str(r.get("Status")).strip().upper() == "CLIENTE" else "Lead",
              "data": r.get("DtCadastro") or None} for r in prospects]

    sk = lambda x: (x.get("nome") or "")
    return {
        "docId": "atual", "date": TODAY.isoformat(), "monthKey": TODAY.strftime("%Y-%m"),
        "ativos": sorted(ativos, key=sk), "vip": sorted(from_ids(vip_ids), key=sk),
        "suspensos": sorted(from_ids(suspended_ids), key=sk), "exClientes": sorted(ex, key=sk),
        "oportunidades": sorted(oportunidades, key=sk), "cancelados": cancelados, "leads": leads,
        "updatedAt": datetime.datetime.utcnow().isoformat() + "Z",
    }


# Cada documento do banco do dashboard aceita no maximo 256 KB, entao as listas sao gravadas
# compactas (linhas em vez de objetos) e divididas em 3 documentos: "base" (ativos, VIP,
# suspensos, cancelados e leads do mes), "exClientes" e "oportunidades".
CLIENT_LIST_COLS = {
    "ativos": ["nome", "contrato", "data"], "vip": ["nome", "contrato", "data"],
    "suspensos": ["nome", "contrato", "data"], "exClientes": ["nome", "contrato", "data"],
    "oportunidades": ["nome", "data"], "cancelados": ["nome", "contrato", "data", "motivo", "valor"],
    "leads": ["nome", "status", "data"],
}


def split_client_lists(doc):
    def pack(key):
        cols = CLIENT_LIST_COLS[key]
        return {"cols": cols, "rows": [[it.get(c) for c in cols] for it in doc[key]]}
    meta = {"date": doc["date"], "monthKey": doc["monthKey"], "updatedAt": doc["updatedAt"]}
    return {
        "base": dict(meta, docId="base", **{k: pack(k) for k in ["ativos", "vip", "suspensos", "cancelados", "leads"]}),
        "exClientes": dict(meta, docId="exClientes", exClientes=pack("exClientes")),
        "oportunidades": dict(meta, docId="oportunidades", oportunidades=pack("oportunidades")),
    }


def build_for_unit(unit):
    dns = unit["dns"]
    token = unit["token"]
    api_get_raw = make_api_get_raw(dns, token)
    API_BASE = f"{API_BASE_ROOT}/api/v2/management"
    API_BASE_V2 = f"{API_BASE_ROOT}/api/v2"
    API_BASE_V3 = f"{API_BASE_ROOT}/api/v3"

    def api_get(base, path, params=None):
        url = f"{base}{path}"
        if params:
            qs = "&".join(f"{k}={v}" for k, v in params.items())
            url = f"{url}?{qs}"
        return api_get_raw(url)

    TODAY = datetime.date.today()
    month_start, month_end = month_bounds(TODAY)
    if TODAY.month == 12:
        next_month_anchor = TODAY.replace(year=TODAY.year + 1, month=1, day=1)
    else:
        next_month_anchor = TODAY.replace(month=TODAY.month + 1, day=1)
    next_month_start, next_month_end = month_bounds(next_month_anchor)

    print(f"[{unit['key']}] Today: {TODAY}, month: {month_start}..{month_end}", file=sys.stderr)

    # Pedido do usuario em 2026-09-29: "Contratos vencendo" mostra ultimo acesso e quantos acessos
    # nos ultimos 30 e 90 dias — guarda todas as datas de entrada (catraca/facial) dos 90 dias.
    entry_dates_by_member = {}
    last_entry_by_member, visits_7d_by_member = get_recent_entries(api_get_raw, days=90, recent_days=7, all_dates_out=entry_dates_by_member)
    _cut30 = (TODAY - datetime.timedelta(days=29)).isoformat()

    def access_stats(idm):
        ds = entry_dates_by_member.get(idm) or set()
        return last_entry_by_member.get(idm), sum(1 for d in ds if d >= _cut30), len(ds)
    print(f"[{unit['key']}] ultima frequencia: {len(last_entry_by_member)} alunos com acesso nos ultimos 90 dias", file=sys.stderr)

    active = load_rows(api_get(API_BASE, "/activeclients"))
    active_ids = set(int(c["IdCliente"]) for c in active if c.get("IdCliente"))
    print(f"[{unit['key']}] active clients (activeclients): {len(active)}", file=sys.stderr)

    member_cache = {}
    all_members_full = json.loads(api_get_raw(f"{API_BASE_V2}/members?take=10000"))
    all_members = [m for m in all_members_full if m.get("status") == "Active"]
    inactive_members = [m for m in all_members_full if m.get("status") == "Inactive"]
    for m in all_members_full:
        member_cache[m["idMember"]] = {"consultor": m.get("nameEmployeeConsultant"), "professor": m.get("nameEmployeeInstructor")}
    print(f"[{unit['key']}] members Active (status=1): {len(all_members)} | Inactive: {len(inactive_members)}", file=sys.stderr)

    t_contracts0 = time.time()
    all_contracts = json.loads(api_get_raw(f"{API_BASE_V3}/membermembership?take=10000&showAggregators=true&showVips=true&showTransfers=true"))
    print(f"[{unit['key']}] membermembership: {len(all_contracts)} contratos em {time.time()-t_contracts0:.0f}s", file=sys.stderr)

    contracts_by_member = {}
    for c in all_contracts:
        c["_start"] = parse_iso_date(c.get("membershipStart"))
        c["_end"] = parse_iso_date(c.get("membershipEnd"))
        c["_cancelDate"] = parse_iso_date(c.get("cancelDate"))
        if c["_cancelDate"] and (c["_end"] is None or c["_cancelDate"] < c["_end"]):
            c["_endEfetivo"] = c["_cancelDate"]
        else:
            c["_endEfetivo"] = c["_end"]
        contracts_by_member.setdefault(c["idMember"], []).append(c)
    for _lst in contracts_by_member.values():
        _lst.sort(key=lambda x: x["_start"] or datetime.date.min)

    def current_contract_for(idm):
        contracts = [c for c in (contracts_by_member.get(idm) or []) if not is_transfer_record(c)]
        if not contracts:
            return None, None
        covering = [c for c in contracts if c["_start"] and c["_end"] and c["_start"] <= TODAY <= c["_end"]]
        if covering:
            c = max(covering, key=lambda x: x["_start"])
            return (c.get("nameMembership") or "").strip(), c["_end"].isoformat()
        started = [c for c in contracts if c["_start"] and c["_start"] <= TODAY]
        if started:
            c = max(started, key=lambda x: x["_start"])
            return (c.get("nameMembership") or "").strip(), c["_end"].isoformat() if c["_end"] else None
        future = [c for c in contracts if c["_start"] and c["_start"] > TODAY]
        if future:
            c = min(future, key=lambda x: x["_start"])
            return (c.get("nameMembership") or "").strip(), c["_end"].isoformat() if c["_end"] else None
        return None, None

    status1_ids = set(m["idMember"] for m in all_members)
    delta_ids = sorted(status1_ids - active_ids)
    vip_count = 0
    suspended_count = 0
    vip_by_contract_type = {}
    vip_ids, suspended_ids = [], []
    for idm in delta_ids:
        try:
            detail = json.loads(api_get_raw(f"{API_BASE_V2}/members/{idm}"))
        except Exception:
            continue
        ms = detail.get("membershipStatus")
        if ms == "Suspended":
            suspended_count += 1
            suspended_ids.append(idm)
        elif ms == "Active":
            vip_count += 1
            vip_ids.append(idm)
            cname = (current_contract_for(idm)[0] or "Outro").strip() or "Outro"
            vip_by_contract_type[cname] = vip_by_contract_type.get(cname, 0) + 1
    print(f"[{unit['key']}] vip={vip_count} suspended={suspended_count} (delta={len(delta_ids)})", file=sys.stderr)

    ex_clients_count = sum(1 for m in inactive_members if m["idMember"] in contracts_by_member)
    print(f"[{unit['key']}] exClients={ex_clients_count} (de {len(inactive_members)} inativos)", file=sys.stderr)

    all_time_prospects = load_rows(api_get(API_BASE, "/prospects", {"dtStart": "2015-01-01", "dtEnd": TODAY.isoformat()}))
    opportunities_count = sum(1 for r in all_time_prospects if str(r.get("Status")).strip().upper() == "PROSPECT")
    print(f"[{unit['key']}] opportunities (historico, nunca convertidos)={opportunities_count} de {len(all_time_prospects)} prospeccoes totais", file=sys.stderr)

    not_renewed_mtd = load_rows(api_get(API_BASE, "/not-renewed", {"dtStart": month_start.isoformat(), "dtEnd": TODAY.isoformat()}))
    prospects = load_rows(api_get(API_BASE, "/prospects", {"dtStart": month_start.isoformat(), "dtEnd": TODAY.isoformat()}))

    cancelled = [r for r in not_renewed_mtd if str(r.get("FlCancelado")).strip().lower() == "true"]
    cancelled_value = sum(parse_value_br(r.get("Valor")) for r in cancelled)
    reasons = {}
    for r in cancelled:
        cat = categorize_reason(r.get("MotivoCancelamento"))
        reasons[cat] = reasons.get(cat, 0) + 1
    internal_plan_changes = reasons.get("Troca interna de plano", 0)

    total_leads = len(prospects)
    converted = [r for r in prospects if str(r.get("Status")).strip().upper() == "CLIENTE"]

    def build_renewal_report(range_start, range_end):
        due = []
        for idm, contracts in contracts_by_member.items():
            for c in contracts:
                if c["_end"] is None or c.get("statusMemberMembership") != 1:
                    continue
                if not (range_start <= c["_end"] <= range_end):
                    continue
                if is_transfer_record(c):
                    continue
                if is_excluded_from_renewals(c.get("nameMembership")):
                    continue
                # Pedido do usuario em 2026-09-29: no FITNESS, troca de plano na renovacao (ex. Fitness
                # Anual DCC -> Fitness Assinatura Plus, 1 dia depois) E renovacao — o contrato antigo
                # continua em "Contratos vencendo" e conta como renovado. Validado contra o relatorio
                # de renovacao do EVO de setembro/2026 (34 contratos, 19 renovados; 17 renovacoes com
                # troca de plano estavam sumindo). A regra de troca de modalidade continua so na
                # Piscina, onde foi validada com o usuario (Joao Vicente, Lunna, Leonardo).
                if unit["key"] == "piscina" and is_internal_plan_swap(c, contracts):
                    continue
                if is_manually_excluded_from_renewals(idm, c):
                    continue
                # Pedido do usuario em 2026-09-29: renovado = existe QUALQUER contrato novo comecando
                # depois do fim deste, mesmo com outro plano. Excecao: "Cortesia 15 dias" nao conta
                # como renovacao — se o contrato seguinte for cortesia, olha-se o contrato pago que
                # vem depois dela (se houver). A propria cortesia continua fora da lista (keyword
                # "cortesia" em _RENEWAL_EXCLUDE_KEYWORDS), entao o que vale e o plano anterior a ela.
                later = [x for x in contracts if x["_start"] and x["_start"] > c["_end"] and not is_transfer_record(x)
                         and "cortesia" not in _strip_accents((x.get("nameMembership") or "").strip().lower())]
                due.append((idm, c, len(later) > 0))

        ids_needed = sorted({idm for idm, _c, _r in due if idm not in member_cache})
        if ids_needed:
            chunk = ",".join(str(i) for i in ids_needed)
            extra_members = json.loads(api_get_raw(f"{API_BASE_V2}/members?idsMembers={chunk}&take=1000"))
            for m in extra_members:
                member_cache[m["idMember"]] = {"consultor": m.get("nameEmployeeConsultant"), "professor": m.get("nameEmployeeInstructor")}

        items = []
        renewed = not_renewed_n = 0
        for idm, c, is_renewed in due:
            e = member_cache.get(idm, {})
            is_future_due = c["_end"] > TODAY
            status = "renovado" if is_renewed else ("a_vencer" if is_future_due else "nao_renovado")
            ult, a30, a90 = access_stats(idm)
            items.append({
                "id": idm, "nome": c.get("name"), "dataFim": c["_end"].strftime("%d/%m/%Y"),
                "plano": (c.get("nameMembership") or "").strip(), "status": status,
                "professor": e.get("professor"), "consultor": e.get("consultor"),
                # Pedido do usuario em 2026-09-29: acessos = dias com entrada na catraca/facial.
                # ultimoAcesso None = nenhum acesso nos ultimos 90 dias.
                "ultimoAcesso": ult, "acessos30d": a30, "acessos90d": a90,
            })
            if is_renewed:
                renewed += 1
            elif status == "nao_renovado":
                # Pedido do usuario em 2026-09-29: "a vencer" nao entra mais em "Nao renovados"
                # (antes o contador mostrava 15 e a lista 14 + 1 a vencer).
                not_renewed_n += 1

        items.sort(key=lambda x: parse_date_br(x["dataFim"]) or datetime.date.max)
        total = len(items)
        pending = sum(1 for it in items if it["status"] == "a_vencer")
        return {
            "total": total, "renewed": renewed, "notRenewed": not_renewed_n, "pending": pending,
            "renewedPct": round(renewed / total * 100, 1) if total else 0,
            "notRenewedPct": round(not_renewed_n / total * 100, 1) if total else 0,
            "items": items,
        }

    renewal_report_this_month = build_renewal_report(month_start, month_end)
    renewal_report_next_month = build_renewal_report(next_month_start, next_month_end)
    print(f"[{unit['key']}] this month: total={renewal_report_this_month['total']} renewed={renewal_report_this_month['renewed']}", file=sys.stderr)

    def build_sales_report(range_start, range_end):
        fetch_end = range_end + datetime.timedelta(days=1)
        sales = json.loads(api_get_raw(
            f"{API_BASE_V2}/sales?dateSaleStart={range_start.isoformat()}&dateSaleEnd={fetch_end.isoformat()}&take=1000&skip=0"
        ))

        ids_needed = sorted({s["idMember"] for s in sales if s.get("idMember") and s["idMember"] not in member_cache})
        if ids_needed:
            chunk = ",".join(str(i) for i in ids_needed)
            extra_members = json.loads(api_get_raw(f"{API_BASE_V2}/members?idsMembers={chunk}&take=1000"))
            for m in extra_members:
                member_cache[m["idMember"]] = {"consultor": m.get("nameEmployeeConsultant"), "professor": m.get("nameEmployeeInstructor")}

        RETORNO_GAP_DAYS = 30

        def sale_type_for(idm, sale_date_d):
            # Pedido do usuario em 2026-09-25 (4a correcao do dia): alem de novo/renovacao/retorno,
            # tambem retorna se uma "renovacao" e ANTECIPADA — cliente comprou o novo contrato antes
            # do mes-calendario em que o contrato anterior realmente venceria (_endEfetivo em mes/ano
            # POSTERIOR ao mes/ano da nova venda). Regra validada nome-a-nome contra a planilha do
            # usuario (15 renovacoes de Piscina em setembro/2026: 8 normais + 7 antecipadas). So se
            # aplica a "renovacao"; None para "novo"/"retorno". Retorna (tipoVenda, renovacaoAntecipada).
            if not idm or not sale_date_d:
                return "novo", None
            # Ajuste manual pontual (pedido do usuario em 2026-09-29) — nao e regra geral.
            _override = _MANUAL_SALE_TYPE_OVERRIDES.get((unit["key"], idm, sale_date_d))
            if _override:
                return _override
            contracts = contracts_by_member.get(idm) or []
            # Pedido do usuario em 2026-09-29: plano CORTESIA e ignorado para classificar a venda —
            # compara sempre com o ultimo plano anterior a cortesia. Caso validado: Diana Cardoso
            # Lima Rodrigues (Fitness, idMember 3232) — Fitness Anual ate 27/02/2024, Cortesia 15 dias
            # 27/07-11/08/2026, nova venda 03/09/2026: antes contava como Renovacao (23 dias apos a
            # cortesia), agora conta como Retorno (2+ anos sem plano pago). Se o cliente so teve
            # cortesia antes, a venda conta como "novo".
            earlier = [c for c in contracts if c["_start"] and c["_start"] < sale_date_d and not is_transfer_record(c)
                       and "cortesia" not in _strip_accents((c.get("nameMembership") or "").strip().lower())]
            if not earlier:
                return "novo", None

            def end_for_gap(c):
                return c["_endEfetivo"] if c["_endEfetivo"] else datetime.date.max

            latest = max(earlier, key=end_for_gap)
            end_date = end_for_gap(latest)
            gap_days = (sale_date_d - end_date).days
            if gap_days > RETORNO_GAP_DAYS:
                return "retorno", None
            is_antecipada = end_date != datetime.date.max and (end_date.year, end_date.month) > (sale_date_d.year, sale_date_d.month)
            return "renovacao", is_antecipada

        items = []
        total_value = 0.0
        for s in sales:
            if s.get("removed"):
                continue
            sale_date = s.get("saleDate")
            try:
                sale_date_d = datetime.datetime.fromisoformat(sale_date).date() if sale_date else None
            except Exception:
                sale_date_d = None
            if sale_date_d and not (range_start <= sale_date_d <= range_end):
                continue
            idm = s.get("idMember")
            e = member_cache.get(idm, {}) if idm else {}
            member = s.get("member") or {}
            nome = " ".join(x for x in [member.get("firstName"), member.get("lastName")] if x).strip() or None
            tipo_venda, renovacao_antecipada = sale_type_for(idm, sale_date_d)
            # Pedido do usuario em 2026-09-29: marca se essa venda e uma TROCA DE CONTRATO detectada
            # automaticamente (cancelamento + venda nova com Adesao no mesmo dia/dia seguinte — ver
            # build_contract_swap_report/troca_sale_ids_mes). Usada abaixo pra excluir essas vendas
            # das contagens "reais" de Novo/Renovacao/Retorno (nao so do valor, como ja era feito).
            is_troca = s["idSale"] in troca_sale_ids_mes
            for it in s.get("saleItens", []):
                val = it.get("saleValue") or 0
                plano = it.get("item")
                items.append({
                    "idSale": s["idSale"], "idMember": idm, "nome": nome,
                    "plano": plano, "categoria": categorize_sale_item(plano), "valor": round(val, 2),
                    "dataVenda": sale_date, "consultor": e.get("consultor"), "tipoVenda": tipo_venda,
                    # Pedido do usuario em 2026-09-25 (4a correcao do dia): so preenchido (True/False)
                    # quando tipoVenda == "renovacao" — ver comentario em sale_type_for().
                    "renovacaoAntecipada": renovacao_antecipada,
                    "trocaContrato": is_troca,
                })
                total_value += val

        items.sort(key=lambda x: x["dataVenda"] or "")

        TIPO_VENDA_LABEL = {"novo": "Novo contrato", "renovacao": "Renovação", "retorno": "Retorno"}
        by_cat = {}
        for it in items:
            cat = it["categoria"]
            # Pedido do usuario em 2026-09-29: "Bike" (ex.: "Bike Mensal") ganha a mesma quebra
            # Novo/Renovação/Renovação Antecipada/Retorno que Fitness e Piscina ja tinham —
            # "tipoVenda" ja era calculado para todo item (sale_type_for() nao filtra por
            # categoria), so nao entrava nessa contagem porque so Fitness/Piscina eram tratados
            # aqui antes.
            if cat in ("Fitness (Anual/Assinatura/Plus)", "Piscina (recorrente)", "Bike"):
                prefixo = "Fitness" if cat.startswith("Fitness") else ("Piscina" if cat.startswith("Piscina") else "Bike")
                # Pedido do usuario em 2026-09-29: venda de troca de contrato ganha categoria propria
                # ("... — Troca de contrato") em vez de entrar em Novo/Renovacao/Retorno — continua
                # visivel em "Vendas por tipo de contrato", so separada das vendas reais.
                if it["trocaContrato"]:
                    tipo_label = "Troca de contrato"
                # Pedido do usuario em 2026-09-25 (4a correcao do dia): categoria propria para
                # "Renovação Antecipada", separada da renovacao normal do mes.
                elif it["tipoVenda"] == "renovacao" and it.get("renovacaoAntecipada"):
                    tipo_label = "Renovação Antecipada"
                else:
                    tipo_label = TIPO_VENDA_LABEL.get(it["tipoVenda"], it["tipoVenda"])
                cat = f"{prefixo} — {tipo_label}"
            agg = by_cat.setdefault(cat, {"count": 0, "valor": 0.0})
            agg["count"] += 1
            agg["valor"] += it["valor"]
        by_category = sorted(
            [{"categoria": k, "count": v["count"], "valor": round(v["valor"], 2)} for k, v in by_cat.items()],
            key=lambda x: -x["count"],
        )

        week_segments = []
        cur = range_start
        while cur <= range_end:
            wk_monday = cur - datetime.timedelta(days=cur.weekday())
            wk_sunday = wk_monday + datetime.timedelta(days=6)
            seg_start = max(wk_monday, range_start)
            seg_end = min(wk_sunday, range_end)
            week_segments.append([seg_start, seg_end])
            cur = seg_end + datetime.timedelta(days=1)

        if len(week_segments) > 1 and (week_segments[0][1] - week_segments[0][0]).days + 1 < 3:
            week_segments[1][0] = week_segments[0][0]
            week_segments.pop(0)
        if len(week_segments) > 1 and (week_segments[-1][1] - week_segments[-1][0]).days + 1 < 3:
            week_segments[-2][1] = week_segments[-1][1]
            week_segments.pop()

        by_week = []
        for seg_start, seg_end in week_segments:
            seg_count = 0
            seg_valor = 0.0
            for it in items:
                dv = it["dataVenda"]
                try:
                    d = datetime.datetime.fromisoformat(dv).date() if dv else None
                except Exception:
                    d = None
                if d is None or not (seg_start <= d <= seg_end):
                    continue
                seg_count += 1
                seg_valor += it["valor"]
            by_week.append({
                "weekStart": seg_start.isoformat(), "weekEnd": seg_end.isoformat(),
                "label": f"{seg_start.strftime('%d/%m')}–{seg_end.strftime('%d/%m')}",
                "count": seg_count, "valor": round(seg_valor, 2),
            })

        # Pedido do usuario em 2026-09-29: as contagens Novo/Renovacao/Renovacao Antecipada/Retorno
        # (e o "Count" total usado no ticket medio) agora EXCLUEM vendas de troca de contrato — antes
        # so o VALOR era descontado (totalValueAjustado), a contagem continuava incluindo as trocas,
        # o que inflava "Renovações" e deflava o ticket medio (numerador ja descontado, denominador
        # nao). Os numeros "como se contasse tudo" (trocas incluidas) ficam disponiveis em campos
        # "*ComTrocas" separados, pra mostrar em letras miudas no dashboard — ver bucket_counts().
        def bucket_counts(lst):
            novos = [it for it in lst if it["tipoVenda"] == "novo"]
            renovacoes = [it for it in lst if it["tipoVenda"] == "renovacao" and not it.get("renovacaoAntecipada")]
            renovacoes_antecipadas = [it for it in lst if it["tipoVenda"] == "renovacao" and it.get("renovacaoAntecipada")]
            retornos = [it for it in lst if it["tipoVenda"] == "retorno"]
            return {
                "count": len(lst), "valor": round(sum(it["valor"] for it in lst), 2),
                "novoCount": len(novos), "novoValue": round(sum(it["valor"] for it in novos), 2),
                "renovacaoCount": len(renovacoes), "renovacaoValue": round(sum(it["valor"] for it in renovacoes), 2),
                "renovacaoAntecipadaCount": len(renovacoes_antecipadas), "renovacaoAntecipadaValue": round(sum(it["valor"] for it in renovacoes_antecipadas), 2),
                "retornoCount": len(retornos), "retornoValue": round(sum(it["valor"] for it in retornos), 2),
            }

        def unit_fields(prefix, categoria, avg_key):
            cats = categoria if isinstance(categoria, (tuple, list)) else (categoria,)
            all_items = [it for it in items if it["categoria"] in cats]
            real_items = [it for it in all_items if not it["trocaContrato"]]
            troca_items = [it for it in all_items if it["trocaContrato"]]
            real = bucket_counts(real_items)
            com_trocas = bucket_counts(all_items)  # real + trocas, mesmo criterio de tipoVenda
            out = {
                f"{prefix}Count": real["count"],
                f"{avg_key}": round(real["valor"] / real["count"], 2) if real["count"] else 0,
                f"{prefix}NovoCount": real["novoCount"], f"{prefix}NovoValue": real["novoValue"],
                f"{prefix}RenovacaoCount": real["renovacaoCount"], f"{prefix}RenovacaoValue": real["renovacaoValue"],
                f"{prefix}RenovacaoAntecipadaCount": real["renovacaoAntecipadaCount"], f"{prefix}RenovacaoAntecipadaValue": real["renovacaoAntecipadaValue"],
                f"{prefix}RetornoCount": real["retornoCount"], f"{prefix}RetornoValue": real["retornoValue"],
                # "ComTrocas" = como ficaria se as vendas de troca de contrato NAO fossem excluidas
                # (fine print no dashboard) — so preenchido quando ha pelo menos 1 troca detectada
                # nessa categoria neste mes, senao e identico ao real (evita ruido visual a toa).
                f"{prefix}CountComTrocas": com_trocas["count"],
                f"{prefix}RenovacaoCountComTrocas": com_trocas["renovacaoCount"], f"{prefix}RenovacaoValueComTrocas": com_trocas["renovacaoValue"],
                f"{prefix}RenovacaoAntecipadaCountComTrocas": com_trocas["renovacaoAntecipadaCount"], f"{prefix}RenovacaoAntecipadaValueComTrocas": com_trocas["renovacaoAntecipadaValue"],
                f"{prefix}NovoCountComTrocas": com_trocas["novoCount"], f"{prefix}NovoValueComTrocas": com_trocas["novoValue"],
                f"{prefix}RetornoCountComTrocas": com_trocas["retornoCount"], f"{prefix}RetornoValueComTrocas": com_trocas["retornoValue"],
                f"{prefix}TrocaContratoCount": len(troca_items), f"{prefix}TrocaContratoValue": round(sum(it["valor"] for it in troca_items), 2),
            }
            return out

        # Pedido do usuario em 2026-09-29: Bike (ex. "Bike Mensal", Paloma) conta junto com o Fitness
        # nos tiles Novos/Renovacoes/Retorno e no numero de vendas da unidade Fitness.
        fitness_fields = unit_fields("fitness", ("Fitness (Anual/Assinatura/Plus)", "Bike"), "avgValue")
        piscina_fields = unit_fields("piscina", "Piscina (recorrente)", "piscinaAvgValue")
        bike_fields = unit_fields("bike", "Bike", "bikeAvgValue")

        out = {
            "total": len(items), "totalValue": round(total_value, 2),
            "byCategory": by_category, "byWeek": by_week, "items": items,
        }
        out.update(fitness_fields)
        out.update(piscina_fields)
        out.update(bike_fields)
        return out

    # Pedido do usuario em 2026-09-29: o relatorio de troca de contrato precisa ser calculado ANTES
    # de build_sales_report() agora, porque este passa a EXCLUIR as vendas de troca da contagem
    # Novo/Renovacao/Retorno (ver "troca_sale_ids_mes" abaixo e "trocaContrato" dentro de
    # build_sales_report) — antes ele so era usado pra descontar o VALOR ("totalValueAjustado"),
    # sem excluir da CONTAGEM, o que inflava fitnessRenovacaoCount/piscinaRenovacaoCount e deflava
    # o ticket medio (numerador ja descontado, denominador nao). Usuario pediu pra ver o numero
    # real nas contagens tambem, com o valor "se contasse tudo" disponivel em letras miudas no
    # dashboard (ver "*ComTrocas" abaixo).
    contract_swap_report = build_contract_swap_report(api_get_raw, api_get, API_BASE, API_BASE_V2, TODAY, month_start)
    print(f"[{unit['key']}] vendas por troca de contrato (ultimos 3 meses): {contract_swap_report['count']} casos, R$ {contract_swap_report['valorTotal']:.2f} (dos quais R$ {contract_swap_report['valorMesCorrente']:.2f} no mes corrente)", file=sys.stderr)
    troca_sale_ids_mes = {
        it["idSaleNova"] for it in contract_swap_report["items"]
        if it["dataVendaNova"] >= month_start.isoformat() and it.get("idSaleNova") is not None
    }
    # Pedido do usuario em 2026-09-29: vendas marcadas MANUALMENTE como troca de contrato (nao
    # detectadas pelo heuristico automatico). Excluidas de Novo/Renovacao/Retorno igual as
    # automaticas, e o valor delas tambem sai do "totalValueAjustado". NAO adicionar ids sem o
    # usuario pedir.
    troca_auto_ids_mes = set(troca_sale_ids_mes)
    troca_sale_ids_mes |= _MANUAL_TROCA_SALE_IDS.get(unit["key"], set())

    sales_report_this_month = build_sales_report(month_start, month_end)
    print(f"[{unit['key']}] sales this month: total={sales_report_this_month['total']} value={sales_report_this_month['totalValue']}", file=sys.stderr)

    loyalty_category = "Fitness (Anual/Assinatura/Plus)" if unit["key"] == "fitness" else "Piscina (recorrente)"
    loyalty_report = build_loyalty_report(loyalty_category, contracts_by_member, TODAY)
    loyalty_report["date"] = TODAY.isoformat()
    print(f"[{unit['key']}] fidelidade: fieis={loyalty_report['fieis']['count']} retornaram={loyalty_report['retornaram']['count']} parcelas={loyalty_report['parcelas']['totalClients']} (media R$ {loyalty_report['parcelas']['avgParcela']:.2f})", file=sys.stderr)

    sales_report_this_month_doc = dict(sales_report_this_month)
    sales_report_this_month_doc["monthKey"] = TODAY.strftime("%Y-%m")
    _manual_troca_items = [it for it in sales_report_this_month["items"]
                           if it["idSale"] in troca_sale_ids_mes and it["idSale"] not in troca_auto_ids_mes]
    _manual_troca_count = len({it["idSale"] for it in _manual_troca_items})
    _manual_troca_value = round(sum(it["valor"] for it in _manual_troca_items), 2)
    sales_report_this_month_doc["countTrocaContratoMes"] = contract_swap_report["countMesCorrente"] + _manual_troca_count
    sales_report_this_month_doc["valorTrocaContratoMes"] = round(contract_swap_report["valorMesCorrente"] + _manual_troca_value, 2)
    sales_report_this_month_doc["totalValueAjustado"] = round(sales_report_this_month["totalValue"] - sales_report_this_month_doc["valorTrocaContratoMes"], 2)
    sales_report_this_month_doc["updatedAt"] = datetime.datetime.utcnow().isoformat() + "Z"

    # Pedido do usuario em 2026-09-27: o usuario notou que "Total de clientes" (e os outros tiles de
    # composicao da base: Clientes ativos, VIP, Suspensos, Ex-clientes, Oportunidades) ficava com o
    # MESMO numero em julho/agosto/setembro ao trocar o filtro de mes — porque esses tiles sempre
    # vinham do snapshot diario mais recente (client_snapshot, abaixo), nunca do relatorio mensal
    # ("clientReports"), entao nao existia historico real por mes (so o numero de HOJE, repetido).
    # Corrigido gravando esses mesmos campos de composicao tambem aqui, no documento mensal
    # "clientReports"/"clientReportsPiscina" (doc_id = mes). Como esse documento SOBRESCREVE so o
    # mes corrente todo dia (mesma regra ja usada pra "cancelledCount"/"newLeads" etc acima) e
    # NUNCA mexe em meses ja fechados, o valor gravado no ultimo dia em que um mes foi "o mes
    # corrente" fica congelado ali pra sempre — funcionando como o fechamento daquele mes. A partir
    # de setembro/2026 (mes em que essa gravacao comecou) pra frente, cada mes fechado vai ter seu
    # proprio numero real de "Total de clientes" etc., em vez de sempre repetir o numero de hoje.
    # Meses anteriores a setembro/2026 continuam sem esse dado (nao da pra reconstruir retroativo,
    # ja explicado ao usuario) — o dashboard cai de volta no snapshot de hoje nesses casos, com nota.
    try:
        client_lists = build_client_lists(TODAY, active, all_members_full, contracts_by_member, current_contract_for,
                                          vip_ids, suspended_ids, all_time_prospects, cancelled, prospects)
    except Exception as e:
        print(f"[{unit['key']}] listas de clientes falharam (resto segue normal): {e}", file=sys.stderr)
        client_lists = None

    client_report_this_month_doc = {
        "monthKey": TODAY.strftime("%Y-%m"),
        "cancelledCount": len(cancelled), "internalPlanChanges": internal_plan_changes,
        "cancelledValue": round(cancelled_value, 2), "cancelReasons": reasons,
        "newLeads": total_leads, "convertedLeads": len(converted),
        "conversionRate": round(len(converted) / total_leads * 100, 1) if total_leads else 0,
        "renewalReport": renewal_report_this_month,
        "activeClients": len(active), "vipClients": vip_count, "vipByContractType": vip_by_contract_type,
        "suspendedClients": suspended_count, "exClients": ex_clients_count, "opportunities": opportunities_count,
        "totalClients": len(active) + suspended_count,
        "updatedAt": datetime.datetime.utcnow().isoformat() + "Z",
    }
    if client_lists is not None:
        client_report_this_month_doc["cancelledList"] = client_lists["cancelados"]
        client_report_this_month_doc["leadsList"] = client_lists["leads"]

    client_snapshot = {
        "date": TODAY.isoformat(), "monthKey": TODAY.strftime("%Y-%m"), "nextMonthKey": next_month_start.strftime("%Y-%m"),
        "activeClients": len(active), "cancelledMTD": len(cancelled), "internalPlanChangesMTD": internal_plan_changes,
        "cancelledValueMTD": round(cancelled_value, 2), "newLeadsMTD": total_leads, "convertedLeadsMTD": len(converted),
        "conversionRateMTD": round(len(converted) / total_leads * 100, 1) if total_leads else 0,
        "cancelReasonsMTD": reasons,
        "renewalReportThisMonth": renewal_report_this_month, "renewalReportNextMonth": renewal_report_next_month,
        "salesReportThisMonth": sales_report_this_month,
        "vipClients": vip_count, "vipByContractType": vip_by_contract_type,
        "suspendedClients": suspended_count, "exClients": ex_clients_count,
        "opportunities": opportunities_count, "totalClients": len(active) + suspended_count,
        "vendasComTrocaDeContrato": contract_swap_report,
        "updatedAt": datetime.datetime.utcnow().isoformat() + "Z",
    }

    workout_snapshot = None
    frequencia_report = None

    if unit["key"] == "piscina":
        freq_list = []
        turma_matched = 0
        RECENT_TURMA_DAYS = 14
        recent_turma_cutoff = (TODAY - datetime.timedelta(days=RECENT_TURMA_DAYS - 1)).isoformat()
        window7_cutoff = (TODAY - datetime.timedelta(days=6)).isoformat()
        lookback_start = TODAY - datetime.timedelta(days=89)
        t_freq0 = time.time()

        for i, m in enumerate(all_members):
            idc = m["idMember"]
            e = member_cache.get(idc, {})
            nome = " ".join(x for x in [m.get("firstName"), m.get("lastName")] if x).strip() or None
            contrato_nome, contrato_fim = current_contract_for(idc)

            try:
                records = get_member_activity_sessions(api_get_raw, idc, lookback_start, TODAY)
            except Exception:
                records = []

            if records:
                turma_matched += 1
                recent = [r for r in records if (r.get("date") or "")[:10] >= recent_turma_cutoff]
                turma = ", ".join(sorted(set(r.get("activitieName") for r in recent if r.get("activitieName")))) or None
                professor = ", ".join(sorted(set((r.get("instructor") or "").strip() for r in recent if (r.get("instructor") or "").strip()))) or None
                within7 = [r for r in records if (r.get("date") or "")[:10] >= window7_cutoff]
                visitas7d = sum(1 for r in within7 if r.get("presenca"))
                presencas = [r["date"][:10] for r in records if r.get("presenca") and r.get("date")]
                ultima_freq = max(presencas) if presencas else None
                fonte = "turma"
            else:
                turma = None
                professor = None
                visitas7d = visits_7d_by_member.get(idc, 0)
                ultima_freq = last_entry_by_member.get(idc)
                fonte = "catraca"

            freq_list.append({
                "idCliente": idc, "nome": nome, "contrato": contrato_nome,
                "vencimentoContrato": contrato_fim,
                "freqContratada": parse_freq_contratada(contrato_nome),
                "visitas7d": visitas7d,
                "bucketVisitas": "0" if visitas7d == 0 else ("1" if visitas7d == 1 else "2+"),
                "ultimaFrequencia": ultima_freq,
                "turma": turma,
                "professor": professor,
                "consultor": e.get("consultor"),
                "fonteFrequencia": fonte,
            })
            if (i + 1) % 50 == 0:
                print(f"[{unit['key']}] frequencia ...{i+1}/{len(all_members)} ({time.time()-t_freq0:.0f}s)", file=sys.stderr)

        freq_list.sort(key=lambda x: (x["nome"] or ""))
        by_freq_contratada = {}
        for c in freq_list:
            by_freq_contratada.setdefault(c["freqContratada"], {"0": 0, "1": 0, "2+": 0})[c["bucketVisitas"]] += 1

        frequencia_report = {
            "date": TODAY.isoformat(), "total": len(freq_list),
            "comMatriculaTurma": turma_matched,
            "byFreqContratada": by_freq_contratada,
            "list": freq_list,
            "updatedAt": datetime.datetime.utcnow().isoformat() + "Z",
        }
        print(f"[{unit['key']}] frequencia: {frequencia_report['total']} alunos, {frequencia_report['comMatriculaTurma']} com turma identificada (fonte=turma)", file=sys.stderr)

        return client_snapshot, workout_snapshot, sales_report_this_month_doc, client_report_this_month_doc, loyalty_report, frequencia_report, client_lists

    def get_client_workout(idc):
        url = f"{API_BASE_V2}/workout/default-client-workout?idClient={idc}&inactive=true"
        return json.loads(api_get_raw(url))

    out = []
    errors = 0
    t0 = time.time()
    for i, m in enumerate(all_members):
        idc = m["idMember"]
        try:
            r = get_client_workout(idc)
        except Exception:
            errors += 1
            continue
        workouts = [w for w in r.get("workouts", []) if not w.get("isDeleted")]
        status = "Sem treino"
        expiry = None
        instructor = None
        if workouts:
            latest = max(workouts, key=lambda w: w.get("creationDate") or "")
            exp = latest.get("expiryDate")
            instructor = latest.get("instructorName")
            if exp:
                exp_date = datetime.datetime.fromisoformat(exp).date()
                expiry = exp_date.isoformat()
                status = "Ativo" if exp_date >= TODAY else "Vencido"
            else:
                status = "Ativo"
        e = member_cache.get(idc, {})
        nome = " ".join(x for x in [m.get("firstName"), m.get("lastName")] if x).strip() or None
        contrato_nome, contrato_fim = current_contract_for(idc)
        out.append({
            "idCliente": idc, "nome": nome, "contratoAtivo": contrato_nome,
            "vencimentoContrato": contrato_fim, "treinoStatus": status,
            # Pedido do usuario em 2026-09-27: o campo "professor" da aba Treinos deve refletir o
            # vinculo de professor ATUALMENTE registrado no cadastro do cliente na EVO (e["professor"],
            # vindo de nameEmployeeInstructor), nao o autor da ultima prescricao de treino (que pode
            # estar desatualizado/expirado ha anos). Prioridade invertida: vinculo atual vem primeiro.
            # Pedido do usuario em 2026-09-28: a carteira de cada professor tem que bater com o
            # relatorio de clientes por professor da EVO — SEM cair de volta no autor do ultimo treino
            # quando o cadastro nao tem professor (isso jogava clientes sem professor na carteira de
            # quem prescreveu o treino). Sem professor no cadastro = "Sem professor" no dashboard.
            "treinoVencimento": expiry, "instrutor": e.get("professor"),
            "consultor": e.get("consultor"), "professorGeral": e.get("professor"),
            "ultimaFrequencia": last_entry_by_member.get(idc),
        })
        if (i + 1) % 100 == 0:
            print(f"[{unit['key']}] ...{i+1}/{len(all_members)} ({time.time()-t0:.0f}s)", file=sys.stderr)

    instr = {}
    for c in out:
        if c["treinoStatus"] == "Ativo" and c["instrutor"]:
            instr[c["instrutor"]] = instr.get(c["instrutor"], 0) + 1
    instr_sorted = dict(sorted(instr.items(), key=lambda x: -x[1]))
    sem_treino = [c for c in out if c["treinoStatus"] == "Sem treino"]
    vencido = [c for c in out if c["treinoStatus"] == "Vencido"]

    EM_7_DIAS = TODAY + datetime.timedelta(days=7)
    vencendo_em_7 = [
        c for c in out
        if c["treinoStatus"] == "Ativo" and c["treinoVencimento"]
        and TODAY <= datetime.date.fromisoformat(c["treinoVencimento"]) <= EM_7_DIAS
    ]

    workout_snapshot = {
        "date": TODAY.isoformat(), "total": len(out), "ativo": sum(1 for o in out if o["treinoStatus"] == "Ativo"),
        "vencido": len(vencido), "semTreino": len(sem_treino), "vencendoEm7Dias": len(vencendo_em_7),
        "instructorBreakdown": instr_sorted,
        "semTreinoList": [{"id": c["idCliente"], "nome": c["nome"], "contrato": c["contratoAtivo"], "vencimentoContrato": c["vencimentoContrato"], "professor": c["professorGeral"], "consultor": c["consultor"], "ultimaFrequencia": c["ultimaFrequencia"]} for c in sem_treino],
        "vencidoList": [{"id": c["idCliente"], "nome": c["nome"], "treinoVencimento": c["treinoVencimento"], "professor": c["instrutor"], "consultor": c["consultor"], "ultimaFrequencia": c["ultimaFrequencia"]} for c in sorted(vencido, key=lambda x: x["treinoVencimento"] or "")],
        "vencendoEm7DiasList": [{"id": c["idCliente"], "nome": c["nome"], "treinoVencimento": c["treinoVencimento"], "professor": c["instrutor"], "consultor": c["consultor"], "ultimaFrequencia": c["ultimaFrequencia"]} for c in sorted(vencendo_em_7, key=lambda x: x["treinoVencimento"] or "")],
        "allClientsList": [
            {
                "id": c["idCliente"], "nome": c["nome"], "treinoStatus": c["treinoStatus"],
                "treinoVencimento": c["treinoVencimento"], "contrato": c["contratoAtivo"],
                "vencimentoContrato": c["vencimentoContrato"],
                "professor": c["professorGeral"], "consultor": c["consultor"],
                "ultimaFrequencia": c["ultimaFrequencia"],
            }
            for c in sorted(out, key=lambda x: (x["nome"] or ""))
        ],
        "updatedAt": datetime.datetime.utcnow().isoformat() + "Z", "errors": errors,
    }

    return client_snapshot, workout_snapshot, sales_report_this_month_doc, client_report_this_month_doc, loyalty_report, frequencia_report, client_lists


if __name__ == "__main__":
    for unit in UNITS:
        snap, workout, sales_doc, client_doc, loyalty_doc, frequencia_doc, lists_doc = build_for_unit(unit)
        suf = unit["suffix"]
        with open(f"snapshot_v3{suf and '_' + suf.lower()}.json", "w") as f:
            json.dump(snap, f, ensure_ascii=False, indent=2)
        if workout is not None:
            with open(f"workout_snapshot_v3{suf and '_' + suf.lower()}.json", "w") as f:
                json.dump(workout, f, ensure_ascii=False, indent=2)
        with open(f"sales_report_month{suf and '_' + suf.lower()}.json", "w") as f:
            json.dump(sales_doc, f, ensure_ascii=False, indent=2)
        with open(f"client_report_month{suf and '_' + suf.lower()}.json", "w") as f:
            json.dump(client_doc, f, ensure_ascii=False, indent=2)
        with open(f"loyalty_report{suf and '_' + suf.lower()}.json", "w") as f:
            json.dump(loyalty_doc, f, ensure_ascii=False, indent=2)
        if frequencia_doc is not None:
            with open(f"frequencia_report{suf and '_' + suf.lower()}.json", "w") as f:
                json.dump(frequencia_doc, f, ensure_ascii=False, indent=2)
        if lists_doc is not None:
            for part_id, part in split_client_lists(lists_doc).items():
                with open(f"client_lists_{part_id}{suf and '_' + suf.lower()}.json", "w") as f:
                    json.dump(part, f, ensure_ascii=False)
        extra_size = len(json.dumps(workout)) if workout is not None else len(json.dumps(frequencia_doc))
        extra_label = "workout" if workout is not None else "frequencia"
        print(f"[{unit['key']}] done. snapshot size={len(json.dumps(snap))} {extra_label} size={extra_size} loyalty size={len(json.dumps(loyalty_doc))}", file=sys.stderr)

