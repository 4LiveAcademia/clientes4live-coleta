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
    # Pedido do usuario em 2026-10-05: janela de "visitas nos ultimos 7 dias" = os 7 dias
    # COMPLETOS anteriores a coleta (ex.: coleta de 05/10 conta de 28/09 a 04/10); o dia da coleta
    # fica de fora porque a coleta roda de madrugada, antes das aulas.
    cutoff = (end - datetime.timedelta(days=recent_days)).date().isoformat()
    cutoff_end = end.date().isoformat()
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
            if cutoff <= d10 < cutoff_end:
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
    if not xlsx_bytes:
        return []
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
    "piscina": {12395, 12396, 12398, 12399, 12400, 12402, 12405, 12406, 12531, 12388, 12421, 12470, 12471},
}
# Pedido do usuario em 2026-09-30 (2a correcao do dia): Maria Clara Ribeiro Meneses (12471) tambem e
# troca de contrato (acima). E vendas que devem ser DESCONSIDERADAS por completo do relatorio de
# vendas (nao aparecem em nenhuma contagem, valor, categoria ou semana), por idSale: Miguel Pacheco
# Oliveira (12375, cobranca de 01/09 00:00 — so a 12388 conta, como troca). NAO adicionar ids sem o
# usuario pedir.
_MANUAL_IGNORE_SALE_IDS = {
    "piscina": {12375},
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
    # Pedido do usuario em 2026-09-30: fora do relatorio de renovacao do EVO de setembro/2026.
    (5826, "2026-08-11"),  # Lunna Maria Melo Farias — Piscina — "Infantil 2x Assinatura" ate 07/09
    (5656, "2025-09-20"),  # Joao Vicente Ribeiro Menezes — Piscina — "INFANTIL 1X ANUAL" ate 16/09
}


def is_manually_excluded_from_renewals(idm, c):
    start = c.get("_start")
    if start is None:
        return False
    return (idm, start.isoformat()) in _MANUAL_RENEWAL_EXCLUSIONS


def categorize_sale_item(name):
    n = (name or "").strip()
    nl = n.lower()
    # Pedido do usuario em 2026-09-30: "Cortesia 15 dias" vira uma categoria propria, com o mesmo
    # nome no Fitness ("CORTESIA 15 DIAS") e na Piscina ("Cortesia 15 dias") — antes cada unidade
    # usava o nome cru do plano. Continua fora de Novo/Renovacao/Retorno (nao e venda de plano pago);
    # o dashboard mostra um tile "Cortesia 15 dias" por unidade com a lista de clientes.
    if "cortesia" in _strip_accents(nl):
        return "Cortesia 15 dias"
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


# Pedido do usuario em 2026-10-06: aba "Permanencia" — quanto tempo cada cliente fica (ou ficou) na
# academia. Para cada cliente com contrato de plano (mesma categoria da Fidelidade: Fitness ou Piscina
# recorrente; transferencias fora): permanencia = dias COM contrato (uniao dos periodos dos contratos,
# sem contar os intervalos sem contrato; contratos cancelados valem ate o cancelamento; ativos contam
# ate hoje) e periodo = do inicio do 1o contrato ate hoje (ativos) ou ate o fim do ultimo (ex-clientes).
PERMANENCIA_MAX_BYTES = 240_000  # cada documento do banco aceita no maximo 256 KB


def build_permanencia(category_filter, contracts_by_member, TODAY):
    ativos, ex = [], []
    for idm, contracts in contracts_by_member.items():
        core = [c for c in contracts if not is_transfer_record(c) and c.get("_start")
                and categorize_sale_item(c.get("nameMembership")) == category_filter]
        if not core:
            continue
        core.sort(key=lambda c: c["_start"])
        if core[0]["_start"] > TODAY:
            continue
        ativo = any(c.get("statusMemberMembership") == 1 and c["_start"] <= TODAY and (c["_end"] is None or TODAY <= c["_end"]) for c in core)
        dias, cur_s, cur_e, ultimo_fim = 0, None, None, None
        for c in core:
            s0 = c["_start"]
            if s0 > TODAY:
                continue
            e0 = min(c["_endEfetivo"] or TODAY, TODAY)
            if e0 < s0:
                continue
            ultimo_fim = e0 if ultimo_fim is None else max(ultimo_fim, e0)
            if cur_s is None:
                cur_s, cur_e = s0, e0
            elif s0 <= cur_e + datetime.timedelta(days=1):
                cur_e = max(cur_e, e0)
            else:
                dias += (cur_e - cur_s).days + 1
                cur_s, cur_e = s0, e0
        if cur_s is not None:
            dias += (cur_e - cur_s).days + 1
        fim = TODAY if ativo else ultimo_fim
        last = max(core, key=lambda c: c["_start"])
        item = {"nome": (last.get("name") or "").strip() or None, "desde": core[0]["_start"].isoformat(),
                "ate": None if ativo else (fim.isoformat() if fim else None), "dias": dias,
                "periodo": ((fim - core[0]["_start"]).days + 1) if fim else dias,
                "plano": (last.get("nameMembership") or "").strip() or None, "contratos": len(core)}
        (ativos if ativo else ex).append(item)

    def resumo(lst):
        if not lst:
            return {"qtd": 0, "mediaDias": 0, "medianaDias": 0}
        ds = sorted(x["dias"] for x in lst)
        n = len(ds)
        med = ds[n // 2] if n % 2 else (ds[n // 2 - 1] + ds[n // 2]) / 2
        return {"qtd": n, "mediaDias": round(sum(ds) / n, 1), "medianaDias": med}

    # ex-clientes por ano de saida (para ver se a permanencia de quem sai esta subindo ou caindo)
    por_ano = {}
    for x in ex:
        if x["ate"]:
            por_ano.setdefault(x["ate"][:4], []).append(x)
    resumo_doc = {
        "categoria": category_filter, "ativos": resumo(ativos), "exClientes": resumo(ex), "todos": resumo(ativos + ex),
        "exPorAnoSaida": {a: resumo(v) for a, v in sorted(por_ano.items())},
    }
    ex.sort(key=lambda x: x["ate"] or "", reverse=True)  # saidas mais recentes primeiro
    return {"resumo": resumo_doc, "ativos": sorted(ativos, key=lambda x: -x["dias"]), "exClientes": ex}


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
                   "valor": round(parse_value_br(r.get("Valor")), 2),
                   "saldo": r.get("_saldo"), "obs": r.get("_obs")} for r in cancelled]

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
    "oportunidades": ["nome", "data"], "cancelados": ["nome", "contrato", "data", "motivo", "valor", "saldo", "obs"],
    "leads": ["nome", "status", "data"],
}


PERMANENCIA_COLS = ["nome", "desde", "ate", "dias", "periodo", "plano", "contratos"]


def split_client_lists(doc):
    def pack(key):
        cols = CLIENT_LIST_COLS[key]
        return {"cols": cols, "rows": [[it.get(c) for c in cols] for it in doc[key]]}
    meta = {"date": doc["date"], "monthKey": doc["monthKey"], "updatedAt": doc["updatedAt"]}
    out = {
        "base": dict(meta, docId="base", **{k: pack(k) for k in ["ativos", "vip", "suspensos", "cancelados", "leads"]}),
        "exClientes": dict(meta, docId="exClientes", exClientes=pack("exClientes")),
        "oportunidades": dict(meta, docId="oportunidades", oportunidades=pack("oportunidades")),
    }
    # Pedido do usuario em 2026-10-06: permanencia — "permanencia" (resumo + ativos) e
    # "permanenciaEx" (ex-clientes). Se passar do limite do documento, a lista de ex-clientes guarda
    # so as saidas mais recentes (o resumo continua calculado com TODOS).
    perm = doc.get("permanencia")
    if perm:
        def ppack(lst):
            return {"cols": PERMANENCIA_COLS, "rows": [[it.get(c) for c in PERMANENCIA_COLS] for it in lst]}
        out["permanencia"] = dict(meta, docId="permanencia", resumo=perm["resumo"], ativos=ppack(perm["ativos"]))
        ex = perm["exClientes"]
        n = len(ex)
        while n > 0 and len(json.dumps(ppack(ex[:n]), ensure_ascii=False).encode()) > PERMANENCIA_MAX_BYTES:
            n = int(n * 0.9)
        out["permanenciaEx"] = dict(meta, docId="permanenciaEx", exClientes=ppack(ex[:n]), total=len(ex), mostrados=n)
    return out


# Pedido do usuario em 2026-09-30: aba "Recebimentos" — contas a receber do mes no EVO
# (/api/v1/receivables, filtro por data de VENCIMENTO dentro do mes), agrupadas por pessoa
# (idMemberPayer), para ver quanto ha para receber, quanto ja entrou e quem esta devendo. Regras:
#  - Recebiveis CANCELADOS ficam fora dos totais (so contados a parte).
#  - Cartao de credito/debito com valor pago preenchido conta como PAGO pelo cliente (o dinheiro vem
#    da adquirente — "Expired"/"open" nesses casos e status de conciliacao, nao inadimplencia).
#  - Pedido do usuario em 2026-09-30 (2a parte): sem separar "atrasado"/"a vencer" — toda cobranca do
#    mes com valor e SEM nada recebido (saldo devedor, boleto, pix... sem valor pago) e INADIMPLENTE.
#  - Pedido do usuario em 2026-09-30 (3a parte): so conta como inadimplente se venceu ATE ONTEM (o dia
#    de hoje ainda nao terminou). Sem pagamento e vencendo hoje ou depois = EM ABERTO (nao e inadimplencia).
#    No dia 1o a coleta tambem refaz o relatorio do MES ANTERIOR (fechamento, com o ultimo dia incluido).
#  - "Ativos sem recebivel no mes": clientes ativos (activeclients) sem nenhuma conta a receber com
#    vencimento no mes — pode ser plano pago a vista (ok) ou cobranca que nao foi gerada (checar).
# Nao grava CPF/e-mail/telefone — so nome, id e valores.
RECEIVABLES_OUT = {}
# Pedido do usuario em 2026-10-01: renovacao vale ate 30 dias depois do vencimento. Nos 30 dias
# seguintes ao fim do mes, a lista de renovacao do mes ANTERIOR e recalculada todo dia e gravada
# (so o campo renewalReport) no documento mensal daquele mes — ver client_report_prev_month*.json.
CLIENT_PREV_OUT = {}
SALES_PREV_OUT = {}
RENOVACAO_PRAZO_DIAS = 30


# Pedido do usuario em 2026-09-30 (4a parte): BOLETO so vira inadimplencia 72h depois do vencimento
# (tempo de compensacao bancaria). Os demais, no dia seguinte ao vencimento.
# Pedido do usuario em 2026-10-01: 72h de tolerancia para TODAS as formas de pagamento (antes so boleto) —
# o cartao recorrente recusado no vencimento tenta de novo nos dias seguintes, e isso gerava falsos
# inadimplentes (ex. 15 no dashboard x 10 no EVO em 01/10/2026).
INAD_TOLERANCIA_DIAS = 3


def _inad_limit(due, payment_type):
    return due + datetime.timedelta(days=INAD_TOLERANCIA_DIAS)


# Pedido do usuario em 2026-09-30 (4a parte): inadimplentes ACUMULADOS — toda cobranca vencida (ate
# ontem; boleto com 72h de carencia) sem nada recebido, dos ULTIMOS 6 MESES (mes corrente + 5 anteriores
# — pedido do usuario em 2026-09-30, 5a parte: nao precisa ir mais para tras), agrupada por pessoa (quantas parcelas em aberto, a mais antiga, total). Gravado junto do
# relatorio do mes corrente (campo inadimplentesAcumulado); no mes seguinte o acumulado segue a partir
# do mes novo, sempre olhando para tras.
def build_delinquency_report(api_get_raw, today, active_ids, names, current_contract_for, status_for=None, access_for=None, months=6):
    # Pedido do usuario em 2026-09-30 (5a parte): tambem traz ultima frequencia, quantas vezes veio nos
    # ultimos 30 dias e a situacao do cliente (Ativo / Suspenso / Inativo).
    y, m = today.year, today.month - (months - 1)
    while m <= 0:
        y, m = y - 1, m + 12
    start = datetime.date(y, m, 1)
    end = today - datetime.timedelta(days=1)
    rows = []
    skip = 0
    while end >= start:
        chunk = json.loads(api_get_raw(f"{API_BASE_ROOT}/api/v1/receivables?dueDateStart={start.isoformat()}"
                                       f"&dueDateEnd={end.isoformat()}&take=1000&skip={skip}"))
        rows += chunk
        if len(chunk) < 1000:
            break
        skip += 1000
    people = {}
    for r in rows:
        st = ((r.get("status") or {}).get("name") or "").lower()
        if st in ("canceled", "received") or float(r.get("ammountPaid") or 0) > 0:
            continue
        pt = (r.get("paymentType") or {}).get("name") if isinstance(r.get("paymentType"), dict) else (r.get("paymentType") or "")
        due = parse_iso_date(r.get("dueDate"))
        if not due or not (_inad_limit(due, pt) < today):
            continue
        valor = float(r.get("ammount") or 0)
        if valor < 1:  # residuos de centavos (ex. arredondamento de parcelas) nao sao inadimplencia
            continue
        idm = r.get("idMemberPayer")
        key = f"m{idm}" if idm else f"n{(r.get('payerName') or '').strip().upper()}"
        p = people.get(key)
        if p is None:
            _ult, _a30 = (access_for(idm)[:2] if (access_for and idm) else (None, None))
            p = people[key] = {"id": idm, "nome": (names.get(idm) if idm else None) or (r.get("payerName") or "").strip() or None,
                               "ativo": bool(idm and idm in active_ids), "status": (status_for(idm) if status_for else None),
                               "contrato": (current_contract_for(idm)[0] if idm else None),
                               "ultimaFrequencia": _ult, "frequencia30d": _a30,
                               "qtd": 0, "total": 0.0, "maisAntiga": None, "parcelas": []}
        p["qtd"] += 1
        p["total"] += valor
        if p["maisAntiga"] is None or due.isoformat() < p["maisAntiga"]:
            p["maisAntiga"] = due.isoformat()
        p["parcelas"].append([r.get("description"), due.isoformat(), round(valor, 2), pt, (today - due).days])
    pessoas = list(people.values())
    for p in pessoas:
        p["total"] = round(p["total"], 2)
        p["parcelas"].sort(key=lambda x: x[1])
    pessoas.sort(key=lambda p: -p["total"])
    return {"ate": (today - datetime.timedelta(days=INAD_TOLERANCIA_DIAS + 1)).isoformat(), "desde": start.isoformat(), "qtdPessoas": len(pessoas),
            "qtdParcelas": sum(p["qtd"] for p in pessoas), "total": round(sum(p["total"] for p in pessoas), 2),
            "pessoas": pessoas}


# Pedido do usuario em 2026-10-01: CAIXA do mes = dinheiro que entra na conta no mes, nao importa quando
# a parcela venceu. Regra da data do dinheiro: cartao (credito, debito, credito on-line) = data prevista de
# credito da parcela (dueDate no EVO — D+30 de cada parcela); demais formas (pix, boleto, dinheiro,
# deposito, saldo devedor...) = data do recebimento (receivingDate). So entra o que foi pago (ammountPaid > 0).
# Bruto = ammountPaid; liquido = ammountPaid - fees (taxa de cartao configurada no EVO).
# Origem: "doMes" (vencimento no proprio mes), "atrasados" (venceu antes, entrou agora), "antecipados"
# (vence depois, entrou agora). "naConta" = data do dinheiro ate hoje; "aCreditar" = cartao ja autorizado
# com credito previsto para mais adiante dentro do mes.
_CARD_TYPES = ("credit card", "debit card", "on-line credit card", "online credit card")


def build_cash_report(api_get_raw, month_start, month_end, today):
    def fetch(param):
        out, skip = [], 0
        while True:
            chunk = json.loads(api_get_raw(f"{API_BASE_ROOT}/api/v1/receivables?{param}Start={month_start.isoformat()}"
                                           f"&{param}End={month_end.isoformat()}&take=1000&skip={skip}"))
            out += chunk
            if len(chunk) < 1000:
                break
            skip += 1000
        return out
    by_id = {}
    for r in fetch("dueDate") + fetch("receivingDate"):
        by_id[r.get("idReceivable") or id(r)] = r
    tot = {"bruto": 0.0, "liquido": 0.0, "taxas": 0.0, "qtd": 0}
    na_conta = {"bruto": 0.0, "liquido": 0.0, "qtd": 0}
    a_creditar = {"bruto": 0.0, "liquido": 0.0, "qtd": 0}
    origem = {k: {"bruto": 0.0, "liquido": 0.0, "qtd": 0} for k in ("doMes", "atrasados", "antecipados")}
    por_forma = {}
    diag = {"cartaoRecebimentoDiferenteDoVencimento": 0, "cartaoTotal": 0}
    for r in by_id.values():
        st = ((r.get("status") or {}).get("name") or "").lower()
        if st == "canceled":
            continue
        pago = float(r.get("ammountPaid") or 0)
        if pago <= 0:
            continue
        pt = (r.get("paymentType") or {}).get("name") if isinstance(r.get("paymentType"), dict) else (r.get("paymentType") or "")
        due = parse_iso_date(r.get("dueDate"))
        rec = parse_iso_date(r.get("receivingDate"))
        is_card = (pt or "").strip().lower() in _CARD_TYPES
        if is_card:
            diag["cartaoTotal"] += 1
            if rec and due and rec != due:
                diag["cartaoRecebimentoDiferenteDoVencimento"] += 1
        money_d = (due or rec) if is_card else (rec or due)
        if not money_d or not (month_start <= money_d <= month_end):
            continue
        taxa = float(r.get("fees") or 0)
        liq = pago - taxa
        tot["bruto"] += pago; tot["liquido"] += liq; tot["taxas"] += taxa; tot["qtd"] += 1
        bucket = na_conta if money_d <= today else a_creditar
        bucket["bruto"] += pago; bucket["liquido"] += liq; bucket["qtd"] += 1
        o = "doMes" if (due and month_start <= due <= month_end) else ("atrasados" if (due and due < month_start) else "antecipados")
        origem[o]["bruto"] += pago; origem[o]["liquido"] += liq; origem[o]["qtd"] += 1
        f = por_forma.setdefault(pt or "Outro", {"qtd": 0, "bruto": 0.0, "liquido": 0.0})
        f["qtd"] += 1; f["bruto"] += pago; f["liquido"] += liq
    r2 = lambda d: {k: (round(v, 2) if isinstance(v, float) else v) for k, v in d.items()}
    return {
        "regra": "cartao = credito previsto da parcela; demais = data do recebimento",
        "ate": today.isoformat(), **r2(tot),
        "naConta": r2(na_conta), "aCreditar": r2(a_creditar),
        "origem": {k: r2(v) for k, v in origem.items()},
        "porForma": {k: r2(v) for k, v in sorted(por_forma.items(), key=lambda kv: -kv[1]["bruto"])},
        "diagnostico": diag,
    }


def build_receivables_report(api_get_raw, month_start, month_end, today, active_ids, names, current_contract_for, status_for=None, access_for=None):
    rows, skip = [], 0
    while True:
        chunk = json.loads(api_get_raw(f"{API_BASE_ROOT}/api/v1/receivables?dueDateStart={month_start.isoformat()}"
                                       f"&dueDateEnd={month_end.isoformat()}&take=1000&skip={skip}"))
        rows += chunk
        if len(chunk) < 1000:
            break
        skip += 1000
    people = {}
    cancel_n, cancel_v = 0, 0.0
    por_forma = {}
    for r in rows:
        st = ((r.get("status") or {}).get("name") or "").lower()
        pt = (r.get("paymentType") or {}).get("name") if isinstance(r.get("paymentType"), dict) else (r.get("paymentType") or "")
        valor = float(r.get("ammount") or 0)
        if st == "canceled":
            cancel_n += 1
            cancel_v += valor
            continue
        pago_v = float(r.get("ammountPaid") or 0)
        due = parse_iso_date(r.get("dueDate"))
        if st == "received" or pago_v > 0:
            sit = "pago"
        elif due and _inad_limit(due, pt) < today:
            sit = "inadimplente"
        else:
            sit = "em_aberto"
        idm = r.get("idMemberPayer")
        key = f"m{idm}" if idm else f"n{(r.get('payerName') or '').strip().upper()}"
        p = people.get(key)
        if p is None:
            p = people[key] = {
                "id": idm, "nome": (names.get(idm) if idm else None) or (r.get("payerName") or "").strip() or None,
                "ativo": bool(idm and idm in active_ids), "status": (status_for(idm) if status_for else None),
                "contrato": (current_contract_for(idm)[0] if idm else None),
                "qtd": 0, "valor": 0.0, "pago": 0.0, "inadimplente": 0.0, "emAberto": 0.0, "parcelas": [],
            }
        p["qtd"] += 1
        p["valor"] += valor
        p["pago" if sit == "pago" else ("inadimplente" if sit == "inadimplente" else "emAberto")] += valor
        p["parcelas"].append([r.get("description"), due.isoformat() if due else None, round(valor, 2), sit, pt])
        f = por_forma.setdefault(pt or "Outro", {"qtd": 0, "valor": 0.0})
        f["qtd"] += 1
        f["valor"] += valor
    pessoas = []
    for p in people.values():
        for k in ("valor", "pago", "inadimplente", "emAberto"):
            p[k] = round(p[k], 2)
        p["situacao"] = "inadimplente" if p["inadimplente"] > 0 else ("em_aberto" if p["emAberto"] > 0 else "pago")
        p["parcelas"].sort(key=lambda x: x[1] or "")
        pessoas.append(p)
    pessoas.sort(key=lambda p: (-p["inadimplente"], (p["nome"] or "")))
    payer_ids = {p["id"] for p in pessoas if p["id"]}
    sem = []
    for idm in sorted(active_ids - payer_ids):
        c, fim = current_contract_for(idm)
        sem.append({"id": idm, "nome": names.get(idm), "contrato": c, "fimContrato": fim, "status": (status_for(idm) if status_for else None)})
    sem.sort(key=lambda x: x["nome"] or "")

    def tot(k):
        return round(sum(p[k] for p in pessoas), 2)
    return {
        "monthKey": month_start.strftime("%Y-%m"), "date": today.isoformat(),
        "totalValor": tot("valor"), "totalPago": tot("pago"), "totalInadimplente": tot("inadimplente"), "totalEmAberto": tot("emAberto"),
        "inadimplenciaAte": (today - datetime.timedelta(days=INAD_TOLERANCIA_DIAS + 1)).isoformat(),
        "qtdRecebiveis": sum(p["qtd"] for p in pessoas), "qtdPessoas": len(pessoas),
        "qtdPessoasInadimplentes": sum(1 for p in pessoas if p["inadimplente"] > 0),
        "canceladosQtd": cancel_n, "canceladosValor": round(cancel_v, 2),
        "porForma": {k: {"qtd": v["qtd"], "valor": round(v["valor"], 2)} for k, v in sorted(por_forma.items(), key=lambda kv: -kv[1]["valor"])},
        "pessoas": pessoas, "ativosSemRecebivel": sem, "qtdAtivosSemRecebivel": len(sem),
        "updatedAt": datetime.datetime.utcnow().isoformat() + "Z",
    }


# =====================================================================================================
# VERSAO 7 (pedido do usuario em 2026-11-11) — SO ACRESCENTA. Tudo acima desta linha e identico ao v6
# e continua gerando os mesmos arquivos. O que vem daqui para baixo gera arquivos NOVOS:
#   parcelas_evo{_piscina}_AAAA-MM_pN.json  -> parcelas com vencimento no mes anterior, corrente e seguinte
#   baixas_evo{_piscina}_pN.json            -> todas as parcelas com baixa no mes corrente (qualquer vencimento)
#   resumo_mes{_piscina}.json               -> classificacao do mes, meta e comissao parcial por consultor
#   resumo_mes_listas{_piscina}.json        -> (so se o resumo passar do limite do documento) listas do resumo
#   resumo_dia{_piscina}.json               -> historico diario (um documento por dia)
#   resumo_mes_consolidado.json / resumo_dia_consolidado.json -> Fitness + Piscina
#   v7_manifest.json                        -> lista de [arquivo, colecao, doc_id] para a rotina gravar
# Somente GET na EVO. Campo que a API nao trouxer fica null e entra em "camposAusentes" — nada e preenchido
# por suposicao. Uma falha aqui nunca afeta os arquivos do v6 (tudo dentro de try/except).
# -----------------------------------------------------------------------------------------------------
# VERSAO 7.1 (pedido do usuario em 2026-11-11) — SO ACRESCENTA ao v7:
#   A. Dias 1 a 5: refaz o resumo, as baixas e as parcelas do MES ANTERIOR com a posicao final (pagamentos
#      ate o ultimo dia) e grava nas mesmas colecoes (resumoMes*, baixasEvo*, parcelasEvo*), doc_id do mes
#      anterior. Arquivos *_prev*. Nesses dias o mes corrente NAO grava as parcelas com vencimento no mes
#      anterior (quem grava e o fechamento). Meta do mes anterior: metas_input_prev.json.
#   B. mes_atual_linhas.json: linhas da aba "Mes atual" da Base de Numeros (uma por unidade + Total),
#      para a rotina acrescentar. Nao vai para o manifesto (nao e gravacao no banco do painel).
#   1. Projecao do mes seguinte movel (todo dia): lancado na EVO + renovacoes previstas (70% +10%, 30% ao
#      ticket de entrada), por unidade e somada, em resumoMes*, resumoDiario* e na aba "Mes atual".
#   2. Dia 1o: meta_congelada.json com a projecao do mes que comeca (Fitness e Piscina); a meta do dia 1o ja
#      usa essa projecao.
#   3. Fora cartao (boleto, Pix...): vencido so depois de 3 dias de tolerancia (igual ao v6).
#   4. Cartao do mes sem autorizacao = recusa: fora da previsao, dentro de vencidoRecusado e da lista de trabalho.
#   5. "renovacao" no resumo do mes: renovacao do mes e renovacao ate hoje.
#   6. Consultor reconhecido so pelo nome inteiro da EVO (V71_CONSULTOR_NOMES_EVO); o run.log lista os nomes.
# =====================================================================================================
V7_FATOR_EVO_BANCO = 1.0          # fator EVO x banco (1,0 por enquanto)
V7_FATOR_META = 1.005             # meta = projecao x fator EVO x banco x 1,005
V7_ALIQ_BASE = 0.005              # 0,5%
V7_ALIQ_META = 0.0075             # 0,75% se a previsao do mes passar da meta
V7_LIMITE_RECUSA_MENSALIDADES = 1.5   # saldo devedor ate 1,5 mensalidade = cobranca recusada; acima = plano em aberto
V7_DIAS_CREDITO_CARTAO = 30       # cartao recorrente: credito ~30 dias depois da cobranca
V7_CONSULTORES = ["Raquel Ribeiro"]   # consultores que recebem comissao hoje; o resto vai para "Sem consultor"
V7_MAX_DOC_BYTES = 240_000        # cada documento do banco do dashboard aceita no maximo 256 KB
V7_MAX_DETALHES_TELEFONE = 150    # limite de GET /members/{id} para achar telefone da lista de trabalho
V7_METAS_ARQ = "metas_input.json" # gerado pela rotina a partir da aba Metas (sem ele, meta fica vazia)

EXTRAS_OUT = {}
EXTRAS_ERR = {}
EXTRAS_PREV_OUT = {}   # v7.1: fechamento do mes anterior (dias 1 a 5)
EXTRAS_PREV_ERR = {}

# campo logico -> chaves candidatas no recebivel (/api/v1/receivables). A primeira que existir e usada.
V7_CAMPOS = {
    "idParcela": ("idReceivable",),
    "idVenda": ("idSale", "saleId"),
    "idContrato": ("idMemberMembership", "idMemberMemberShip"),
    "numeroParcela": ("currentInstallment", "installment", "installmentNumber", "numberInstallment"),
    "totalParcelas": ("totalInstallments", "totalInstallment", "installments", "numberOfInstallments"),
    "descricao": ("description",),
    "vencimento": ("dueDate",),
    "dataBaixa": ("receivingDate",),
    "valor": ("ammount", "amount"),
    "valorPago": ("ammountPaid", "amountPaid"),
    "formaPagamento": ("paymentType",),
    "situacao": ("status",),
    "dataCancelamento": ("cancellationDate", "cancelDate", "canceledDate", "dateCancellation"),
    "motivoCancelamento": ("cancellationReason", "reasonCancellation", "cancellationDescription", "cancelReason"),
    "dataCobranca": ("registrationDate", "chargeDate"),
    "pagamentos": ("payments", "receivablePayments", "paymentHistory", "receivingHistory"),
    "origemSaldoDevedor": ("balanceOrigin", "originBalance", "debitBalanceOrigin", "balanceType"),
    "pagante": ("idMemberPayer",),
}
V7_CAMPOS_MEMBRO = {
    "consultor": ("nameEmployeeConsultant",),
    "telefone": ("cellphone", "cellPhone", "phone", "mobile", "telephone", "contacts"),
}


def _v7_norm(s):
    return " ".join(_strip_accents(str(s or "")).lower().split())


def _v7_get(r, campo, presenca=None):
    for k in V7_CAMPOS[campo]:
        if k in r:
            if presenca is not None:
                presenca.setdefault(campo, k)
            return r.get(k)
    return None


def _v7_nome(v):
    if isinstance(v, dict):
        return v.get("name") or v.get("description")
    return v


def _v7_num(v):
    try:
        return round(float(v), 2) if v is not None and v != "" else None
    except Exception:
        return None


def _v7_date(v):
    d = parse_iso_date(v) if v else None
    return d.isoformat() if d else None


def _v7_pagamentos(v):
    if not isinstance(v, list):
        return None
    out = []
    for p in v[:10]:
        if isinstance(p, dict):
            out.append({k: x for k, x in p.items() if isinstance(x, (str, int, float, bool)) or x is None})
        else:
            out.append(p)
    return out


def _v7_fetch_receivables(api_get_raw, param, d0, d1):
    out, skip = [], 0
    while True:
        chunk = json.loads(api_get_raw(f"{API_BASE_ROOT}/api/v1/receivables?{param}Start={d0.isoformat()}"
                                       f"&{param}End={d1.isoformat()}&take=1000&skip={skip}"))
        out += chunk
        if len(chunk) < 1000:
            break
        skip += 1000
    return out


def _v7_is_saldo_devedor(forma, descricao):
    f, d = _v7_norm(forma), _v7_norm(descricao)
    return ("saldo" in f or "balance" in f or "devedor" in f) or d.startswith("saldo devedor")


def _v7_is_card(forma):
    f = _v7_norm(forma)
    return f in _CARD_TYPES or "credit" in f or "credito" in f or "cartao" in f


def _v7_phone(m):
    if not isinstance(m, dict):
        return None
    for k in ("cellphone", "cellPhone", "mobile", "phone", "telephone"):
        if m.get(k):
            return str(m.get(k)).strip()
    cs = m.get("contacts")
    if isinstance(cs, list):
        prefer, other = None, None
        for c in cs:
            if not isinstance(c, dict):
                continue
            t = _v7_norm(c.get("contactType") or c.get("type") or c.get("description_type") or "")
            val = c.get("description") or c.get("value") or c.get("contact")
            if not val or "@" in str(val):
                continue
            if any(x in t for x in ("cel", "mobile", "whats")):
                prefer = prefer or str(val).strip()
            elif any(x in t for x in ("tel", "phone", "fone")):
                other = other or str(val).strip()
        return prefer or other
    return None


# v7.1 (item 6): consultor reconhecido so pelo NOME INTEIRO que a EVO traz (sem diferenca de maiusculas,
# acentos e espacos), para nao confundir "Raquel Ribeiro" com "Raquel Monique da Cruz Cardoso".
# Chave = nome usado na comissao; valor = nomes exatos aceitos na EVO (nameEmployeeConsultant).
# CONFERIR na 1a execucao com a linha "[v7.1] consultores na EVO" do run.log e, se o nome da EVO for outro,
# trocar aqui (e subir nova versao).
V71_CONSULTOR_NOMES_EVO = {"Raquel Ribeiro": ["Raquel Ribeiro"]}
V71_TICKET_ENTRADA = {"fitness": 458.44, "piscina": 430.23}   # reposicao ao ticket de entrada
V71_RENOVA_PCT = 0.70        # 70% renovam...
V71_RENOVA_REAJUSTE = 0.10   # ...com +10% sobre o valor atual; os outros 30% sao repostos ao ticket de entrada
_V71_LOG_CONSULTORES = set()


def _v7_consultor_canonico(nome_evo):
    n = _v7_norm(nome_evo)
    if not n:
        return None
    for c in V7_CONSULTORES:
        if n in {_v7_norm(x) for x in V71_CONSULTOR_NOMES_EVO.get(c, [c])}:
            return c
    return None


def _v71_log_consultores(key, membros):
    if key in _V71_LOG_CONSULTORES:
        return
    _V71_LOG_CONSULTORES.add(key)
    cont = {}
    for m in membros.values():
        nm = (m.get("nameEmployeeConsultant") or "").strip()
        if nm:
            cont[nm] = cont.get(nm, 0) + 1
    raq = sorted(n for n in cont if "raquel" in _v7_norm(n))
    print(f"[v7.1] consultores na EVO ({key}): " + "; ".join(f"{n} ({q})" for n, q in sorted(cont.items(), key=lambda kv: -kv[1])[:20]), file=sys.stderr)
    print(f"[v7.1] nomes com 'Raquel' na EVO ({key}): {raq or 'nenhum'} | reconhecidos como consultor: "
          f"{[n for n in raq if _v7_consultor_canonico(n)] or 'NENHUM - conferir V71_CONSULTOR_NOMES_EVO'}", file=sys.stderr)


def _v71_renovacoes_previstas(key, renovacao, mensalidade):
    """Contratos que vencem no mes e ainda nao renovaram: 70% x valor atual x 1,10 + 30% x ticket de entrada."""
    ticket = V71_TICKET_ENTRADA.get(key)
    out = {"qtdContratos": 0, "qtdSemValorAtual": 0, "valorAtualSoma": 0.0, "total": 0.0, "ticketEntrada": ticket,
           "pctRenovam": V71_RENOVA_PCT, "reajuste": V71_RENOVA_REAJUSTE, "itens": []}
    if not renovacao:
        out["aviso"] = "sem relatorio de renovacao do mes: renovacoes previstas = 0"
        return out
    for it in renovacao.get("items") or []:
        if it.get("status") == "renovado":
            continue
        v, fonte = mensalidade(it.get("id")) if it.get("id") is not None else (None, None)
        if v is None:
            out["qtdSemValorAtual"] += 1
            v, fonte = ticket, "sem valor atual: usado o ticket de entrada"
        prev = V71_RENOVA_PCT * v * (1 + V71_RENOVA_REAJUSTE) + (1 - V71_RENOVA_PCT) * ticket
        out["qtdContratos"] += 1
        out["valorAtualSoma"] += v
        out["total"] += prev
        out["itens"].append([it.get("id"), it.get("nome"), it.get("dataFim"), it.get("status"), round(v, 2), fonte, round(prev, 2)])
    out["valorAtualSoma"] = round(out["valorAtualSoma"], 2)
    out["total"] = round(out["total"], 2)
    return out


def _v71_renovacao_ate(renovacao, ref):
    if not renovacao:
        return None
    items = renovacao.get("items") or []
    venc = [it for it in items if (parse_date_br(it.get("dataFim")) or datetime.date.max) <= ref]
    ren = sum(1 for it in venc if it.get("status") == "renovado")
    return {"doMes": {"contratos": renovacao.get("total"), "renovados": renovacao.get("renewed"), "pct": renovacao.get("renewedPct")},
            "ateHoje": {"data": ref.isoformat(), "contratosVencidos": len(venc), "renovados": ren,
                        "pct": round(ren / len(venc) * 100, 1) if venc else None},
            "regra": "ate hoje = renovados / contratos que ja venceram ate a data (mesma regra de renovado do relatorio do mes)"}


def _v7_split_rows(meta, doc_prefix, cols, rows):
    """Divide uma lista compacta (cols/rows) em partes que cabem num documento."""
    partes, atual, tam = [], [], 0
    base = len(json.dumps(meta, ensure_ascii=False).encode()) + 600
    for r in rows:
        t = len(json.dumps(r, ensure_ascii=False).encode()) + 1
        if atual and base + tam + t > V7_MAX_DOC_BYTES:
            partes.append(atual)
            atual, tam = [], 0
        atual.append(r)
        tam += t
    if atual or not partes:
        partes.append(atual)
    n = len(partes)
    return [dict(meta, docId=f"{doc_prefix}-p{i+1}", parte=i + 1, partes=n, cols=cols, rows=p) for i, p in enumerate(partes)]


V7_PARCELA_COLS = ["idParcela", "idContrato", "idVenda", "numeroParcela", "totalParcelas", "inicioContrato", "fimContrato",
                   "plano", "descricao", "vencimento", "dataBaixa", "valor", "valorPago", "forma", "situacao",
                   "dataCancelamento", "motivoCancelamento", "dataCobranca", "idCliente", "cliente", "consultorEvo",
                   "classe", "pagamentos"]


def build_month_extras(unit, api_get_raw, api_base_v2, TODAY, month_start, month_end, next_month_start, next_month_end,
                       all_members_full, all_contracts, class_ref=None, buscar_telefone=True, renovacao=None):
    # v7.1: class_ref = data usada para separar "vencido" de "a vencer" (padrao: hoje). Para o fechamento do
    # mes anterior a rotina passa o 1o dia do mes corrente: tudo que venceu ate o ultimo dia ja conta como
    # vencido, e o que foi pago ate o ultimo dia como garantido (posicao final do mes).
    # v7.1: renovacao = relatorio de renovacao do mes calculado (renewalReport do v6), usado na projecao do
    # mes seguinte (renovacoes previstas) e na "renovacao ate hoje".
    key = unit["key"]
    prev_end = month_start - datetime.timedelta(days=1)
    prev_start = prev_end.replace(day=1)
    presenca = {}
    avisos = []

    # ---- 1. recebiveis: vencimento no mes anterior, corrente e seguinte + baixa no mes corrente ----
    raw_venc = _v7_fetch_receivables(api_get_raw, "dueDate", prev_start, next_month_end)
    raw_baixa = _v7_fetch_receivables(api_get_raw, "receivingDate", month_start, month_end)
    print(f"[{key}] v7: {len(raw_venc)} parcelas com vencimento {prev_start}..{next_month_end}, {len(raw_baixa)} com baixa no mes", file=sys.stderr)

    # mesma parcela (mesmo id) aparecendo com baixas diferentes = conferir. Nunca por valor/data parecidos.
    baixas_por_id = {}
    for r in raw_venc + raw_baixa:
        idr = r.get("idReceivable")
        if idr is None:
            continue
        b = (_v7_date(r.get("receivingDate")), _v7_num(r.get("ammountPaid")))
        if b[0] or (b[1] or 0) > 0:
            baixas_por_id.setdefault(idr, set()).add(b)
    by_id = {}
    sem_id = []
    for r in raw_venc + raw_baixa:
        idr = r.get("idReceivable")
        if idr is None:
            sem_id.append(r)
        else:
            by_id[idr] = r
    if sem_id:
        avisos.append(f"{len(sem_id)} recebiveis sem idReceivable (mantidos, sem deduplicar)")
    todos = list(by_id.values()) + sem_id

    # ---- 2. clientes (pagante = idMemberPayer = idMember), consultor e telefone ----
    membros = {m["idMember"]: m for m in all_members_full if m.get("idMember") is not None}
    pagantes = {r.get("idMemberPayer") for r in todos if r.get("idMemberPayer")}
    faltam = sorted(i for i in pagantes if i not in membros)
    for i in range(0, len(faltam), 100):
        try:
            ids = ",".join(str(x) for x in faltam[i:i + 100])
            for m in json.loads(api_get_raw(f"{api_base_v2}/members?idsMembers={ids}&take=1000")):
                if m.get("idMember") is not None:
                    membros[m["idMember"]] = m
        except Exception as e:
            avisos.append(f"falha ao buscar {len(faltam[i:i+100])} pagantes fora da lista de membros: {e}")
    pres_membro = {}
    for m in membros.values():
        for campo, ks in V7_CAMPOS_MEMBRO.items():
            for k in ks:
                if k in m:
                    pres_membro.setdefault(campo, k)

    def nome_de(idm, r=None):
        m = membros.get(idm)
        n = " ".join(x for x in [m.get("firstName"), m.get("lastName")] if x).strip() if m else ""
        return n or ((r or {}).get("payerName") or "").strip() or None

    def consultor_evo(idm):
        m = membros.get(idm)
        return (m.get("nameEmployeeConsultant") or None) if m else None

    # ---- 3. contratos (ligacao so por id: idContrato no recebivel, ou idSale -> contrato) ----
    contrato_por_id = {c.get("idMemberMemberShip"): c for c in all_contracts if c.get("idMemberMemberShip") is not None}
    contratos_por_venda = {}
    for c in all_contracts:
        if c.get("idSale") is not None:
            contratos_por_venda.setdefault(c["idSale"], []).append(c)

    def contrato_de(id_contrato, id_venda):
        if id_contrato is not None and id_contrato in contrato_por_id:
            return contrato_por_id[id_contrato], id_contrato
        cs = contratos_por_venda.get(id_venda) if id_venda is not None else None
        if cs and len(cs) == 1:  # venda com mais de um contrato: nao escolhe por suposicao
            return cs[0], cs[0].get("idMemberMemberShip")
        return None, id_contrato

    # ---- 4. linha de cada parcela ----
    linhas = {}
    for r in todos:
        idv = _v7_get(r, "idVenda", presenca)
        idc_raw = _v7_get(r, "idContrato", presenca)
        c, idc = contrato_de(idc_raw, idv)
        pag = _v7_get(r, "pagante", presenca)
        forma = _v7_nome(_v7_get(r, "formaPagamento", presenca))
        sit = _v7_nome(_v7_get(r, "situacao", presenca))
        row = {
            "idParcela": _v7_get(r, "idParcela", presenca), "idContrato": idc, "idVenda": idv,
            "numeroParcela": _v7_get(r, "numeroParcela", presenca), "totalParcelas": _v7_get(r, "totalParcelas", presenca),
            "inicioContrato": c["_start"].isoformat() if c and c.get("_start") else None,
            "fimContrato": c["_end"].isoformat() if c and c.get("_end") else None,
            "plano": ((c.get("nameMembership") or "").strip() or None) if c else None,
            "descricao": _v7_get(r, "descricao", presenca), "vencimento": _v7_date(_v7_get(r, "vencimento", presenca)),
            "dataBaixa": _v7_date(_v7_get(r, "dataBaixa", presenca)),
            "valor": _v7_num(_v7_get(r, "valor", presenca)), "valorPago": _v7_num(_v7_get(r, "valorPago", presenca)),
            "forma": forma, "situacao": sit,
            "dataCancelamento": _v7_date(_v7_get(r, "dataCancelamento", presenca)),
            "motivoCancelamento": _v7_get(r, "motivoCancelamento", presenca),
            "dataCobranca": _v7_date(_v7_get(r, "dataCobranca", presenca)),
            "idCliente": pag, "cliente": nome_de(pag, r), "consultorEvo": consultor_evo(pag),
            "pagamentos": _v7_pagamentos(_v7_get(r, "pagamentos", presenca)),
            "_origemSaldo": _v7_get(r, "origemSaldoDevedor", presenca),
        }
        linhas[row["idParcela"] if row["idParcela"] is not None else id(r)] = row

    # mensalidade do cliente (para a regra de 1,5x): valor da parcela mais recente que NAO e saldo devedor,
    # nem cancelada, dentro da janela coletada; senao, valor do contrato atual / meses do contrato.
    mens = {}
    for row in sorted(linhas.values(), key=lambda x: x["vencimento"] or ""):
        if row["idCliente"] and _v7_norm(row["situacao"]) != "canceled" and not _v7_is_saldo_devedor(row["forma"], row["descricao"]) and (row["valor"] or 0) > 0:
            mens[row["idCliente"]] = (row["valor"], "ultima parcela que nao e saldo devedor")

    def mensalidade(idm):
        if idm in mens:
            return mens[idm]
        cs = [c for c in all_contracts if c.get("idMember") == idm and c.get("_start") and c.get("_end") and c["_start"] <= TODAY]
        if cs:
            c = max(cs, key=lambda x: x["_start"])
            meses = max(1, round((c["_end"] - c["_start"]).days / 30.4))
            if (c.get("saleValue") or 0) > 0:
                return round(c["saleValue"] / meses, 2), f"contrato atual / {meses} meses"
        return None, None

    # ---- 5. classificacao do mes corrente ----
    ms, me, hoje = month_start.isoformat(), month_end.isoformat(), (class_ref or TODAY).isoformat()
    G = {"total": 0.0, "qtd": 0, "cartaoAutorizado": 0.0, "outrasFormasComBaixa": 0.0, "vencimentoPassado": 0.0, "vencimentoFuturo": 0.0}
    A = {"total": 0.0, "qtd": 0, "porForma": {}}
    CSA = {"total": 0.0, "qtd": 0, "itens": []}
    venc_itens, recusas, plano_aberto, saldo_sem = [], [], [], []
    card_itens = []  # v7.1: cartao com vencimento no mes sem autorizacao
    base_baixa = {}  # idCliente -> valor com baixa no mes (garantido + recuperado + antecipado)

    def soma_base(idm, v, r):
        k = idm if idm else f"n:{r.get('cliente') or ''}"
        b = base_baixa.setdefault(k, {"idCliente": idm, "cliente": r.get("cliente"), "consultorEvo": r.get("consultorEvo"), "valor": 0.0})
        b["valor"] += v

    for row in linhas.values():
        st = _v7_norm(row["situacao"])
        if st == "canceled":
            row["classe"] = "cancelada"
            continue
        v_due, pago_v, val = row["vencimento"], (row["valorPago"] or 0), (row["valor"] or 0)
        pago = pago_v > 0 or st == "received"
        saldo = _v7_is_saldo_devedor(row["forma"], row["descricao"])
        if saldo and not pago:
            mval, mfonte = mensalidade(row["idCliente"])
            origem_api = row["_origemSaldo"]
            item = {"idCliente": row["idCliente"], "cliente": row["cliente"], "descricao": row["descricao"],
                    "vencimento": v_due, "dataCobranca": row["dataCobranca"], "valor": round(val, 2),
                    "mensalidade": mval, "mensalidadeFonte": mfonte, "origemNaEvo": origem_api,
                    "consultorEvo": row["consultorEvo"], "idParcela": row["idParcela"]}
            if mval is None:
                row["classe"] = "saldo_sem_classificacao"
                saldo_sem.append(item)
            elif val <= V7_LIMITE_RECUSA_MENSALIDADES * mval:
                base = parse_iso_date(row["dataCobranca"] or v_due)
                item["mesDoDinheiro"] = (base + datetime.timedelta(days=V7_DIAS_CREDITO_CARTAO)).strftime("%Y-%m") if base else None
                row["classe"] = "recusa"
                recusas.append(item)
            else:
                row["classe"] = "plano_em_aberto"
                plano_aberto.append(item)
            continue
        if not v_due:
            row["classe"] = "sem_vencimento"
            continue
        if ms <= v_due <= me:
            if pago:
                v = pago_v if pago_v > 0 else val
                G["total"] += v; G["qtd"] += 1
                G["cartaoAutorizado" if _v7_is_card(row["forma"]) else "outrasFormasComBaixa"] += v
                G["vencimentoPassado" if v_due < hoje else "vencimentoFuturo"] += v
                row["classe"] = "garantido"
                soma_base(row["idCliente"], v, row)
            elif _v7_is_card(row["forma"]):
                # v7.1 (item 4): todo cartao que vence no mes ja foi cobrado no mes anterior; sem autorizacao = recusa.
                # Fica fora da previsao e entra na lista de trabalho do time (vencido, nao pago e recusado).
                row["classe"] = "cartao_sem_autorizacao"
                CSA["total"] += val; CSA["qtd"] += 1
                CSA["itens"].append([row["idCliente"], row["cliente"], row["descricao"], v_due, round(val, 2), row["forma"]])
                card_itens.append({"tipo": "cartão sem autorização (recusa)", "idCliente": row["idCliente"], "cliente": row["cliente"],
                                   "descricao": row["descricao"], "vencimento": v_due, "valor": round(val, 2),
                                   "forma": row["forma"], "consultorEvo": row["consultorEvo"], "mesDoDinheiro": v_due[:7],
                                   "diasVencido": max(0, (TODAY - parse_iso_date(v_due)).days), "idParcela": row["idParcela"]})
            elif _inad_limit(parse_iso_date(v_due), row["forma"]).isoformat() >= hoje:
                # v7.1 (item 3): boleto, Pix e demais formas fora cartao so viram vencido depois de 3 dias (igual ao v6)
                row["classe"] = "em_aberto"
                A["total"] += val; A["qtd"] += 1
                f = A["porForma"].setdefault(row["forma"] or "Outro", {"qtd": 0, "valor": 0.0})
                f["qtd"] += 1; f["valor"] += val
            else:
                row["classe"] = "vencido_nao_pago"
                venc_itens.append({"tipo": "parcela vencida", "idCliente": row["idCliente"], "cliente": row["cliente"],
                                   "descricao": row["descricao"], "vencimento": v_due, "valor": round(val, 2),
                                   "forma": row["forma"], "consultorEvo": row["consultorEvo"], "mesDoDinheiro": v_due[:7],
                                   "diasVencido": (TODAY - parse_iso_date(v_due)).days, "idParcela": row["idParcela"]})
        elif v_due < ms:
            row["classe"] = "mes_anterior_pago" if pago else "mes_anterior_aberto"
        else:
            row["classe"] = "mes_seguinte_pago" if pago else "mes_seguinte_aberto"

    # ---- 6. baixas do mes (qualquer vencimento): carteira do mes, recuperado, antecipado ----
    R = {"total": 0.0, "qtd": 0}
    AN = {"total": 0.0, "qtd": 0}
    CM = {"total": 0.0, "qtd": 0}
    baixa_rows = []
    for r in raw_baixa:
        row = linhas.get(r.get("idReceivable")) if r.get("idReceivable") is not None else None
        if row is None or _v7_norm(row["situacao"]) == "canceled":
            continue
        pago_v = row["valorPago"] or 0
        if pago_v <= 0 and _v7_norm(row["situacao"]) != "received":
            continue
        if not (row["dataBaixa"] and ms <= row["dataBaixa"] <= me):
            continue
        v = pago_v if pago_v > 0 else (row["valor"] or 0)
        due = row["vencimento"]
        if due and ms <= due <= me:
            origem = "carteira_do_mes"; CM["total"] += v; CM["qtd"] += 1
        elif due and due < ms:
            origem = "recuperado"; R["total"] += v; R["qtd"] += 1
            soma_base(row["idCliente"], v, row)
        elif due and due > me:
            origem = "antecipado"; AN["total"] += v; AN["qtd"] += 1
            soma_base(row["idCliente"], v, row)
        else:
            origem = "sem_vencimento"
        baixa_rows.append([row["idParcela"], row["idCliente"], row["cliente"], row["consultorEvo"], row["descricao"], due,
                           row["dataBaixa"], row["valor"], row["valorPago"], row["forma"], row["situacao"], origem,
                           row["idContrato"], row["idVenda"]])
    baixa_rows.sort(key=lambda x: (x[6] or "", x[2] or ""))

    # ---- 7. conferir: mesma parcela (mesmo id) com duas baixas ----
    conferir = []
    for idr, bs in baixas_por_id.items():
        if len(bs) > 1:
            row = linhas.get(idr) or {}
            conferir.append({"idParcela": idr, "cliente": row.get("cliente"), "idCliente": row.get("idCliente"),
                             "baixas": sorted([list(b) for b in bs], key=lambda b: b[0] or ""), "fonte": "mesmo id com baixas diferentes"})
    for row in linhas.values():
        pg = row.get("pagamentos") or []
        datas = [p for p in pg if isinstance(p, dict)]
        if len(datas) >= 2:
            conferir.append({"idParcela": row["idParcela"], "cliente": row["cliente"], "idCliente": row["idCliente"],
                             "baixas": datas, "fonte": "lista de pagamentos da parcela"})

    # ---- 8. telefone so para a lista de trabalho ----
    trab = venc_itens + card_itens + [dict(x, tipo="cobrança recusada") for x in recusas]
    precisa = sorted({x["idCliente"] for x in trab if x.get("idCliente") and not _v7_phone(membros.get(x["idCliente"]))})
    detalhes = 0
    tel_cache = {}
    if buscar_telefone and ("telefone" in pres_membro or precisa):
        for idm in precisa[:V7_MAX_DETALHES_TELEFONE]:
            try:
                d = json.loads(api_get_raw(f"{api_base_v2}/members/{idm}"))
                detalhes += 1
                tel_cache[idm] = _v7_phone(d)
                if _v7_phone(d):
                    pres_membro.setdefault("telefone", "members/{id}")
            except Exception:
                tel_cache[idm] = None
        if len(precisa) > V7_MAX_DETALHES_TELEFONE:
            avisos.append(f"telefone buscado so para {V7_MAX_DETALHES_TELEFONE} de {len(precisa)} clientes da lista")
    for x in trab:
        idm = x.get("idCliente")
        x["telefone"] = _v7_phone(membros.get(idm)) or tel_cache.get(idm)
        x["unidade"] = unit["label"]
    trab.sort(key=lambda x: (x.get("mesDoDinheiro") or "", x.get("vencimento") or "", x.get("cliente") or ""))

    cur_mk = month_start.strftime("%Y-%m")  # v7.1: mes do calculo (igual ao de hoje, salvo no fechamento do mes anterior)
    rec_mes = [x for x in recusas if x.get("mesDoDinheiro") == cur_mk]
    rec_outros = {}
    for x in recusas:
        if x.get("mesDoDinheiro") != cur_mk:
            o = rec_outros.setdefault(x.get("mesDoDinheiro") or "sem data", {"qtd": 0, "valor": 0.0})
            o["qtd"] += 1; o["valor"] += x["valor"]
    v_parc = round(sum(x["valor"] for x in venc_itens), 2)
    v_rec = round(sum(x["valor"] for x in rec_mes), 2)
    v_card = round(sum(x["valor"] for x in card_itens), 2)

    # ---- 9. campos ausentes (existencia da chave em pelo menos um registro) ----
    ausentes = []
    if todos:
        for campo in V7_CAMPOS:
            if campo not in presenca:
                ausentes.append(f"recebivel.{campo}")
    for campo in V7_CAMPOS_MEMBRO:
        if campo not in pres_membro:
            ausentes.append(f"membro.{campo}")
    if "origemSaldoDevedor" not in presenca:
        avisos.append("a API nao informou a origem do saldo devedor: usada a regra de 1,5 mensalidade")
    if "pagamentos" not in presenca:
        avisos.append("a API nao trouxe pagamentos por parcela: 'conferir' usa so o mesmo id com baixas diferentes")

    r2 = lambda x: round(x, 2)
    for d in (G, A, R, AN, CM, CSA):
        d["total"] = r2(d["total"])
    for k in ("cartaoAutorizado", "outrasFormasComBaixa", "vencimentoPassado", "vencimentoFuturo"):
        G[k] = r2(G[k])
    A["porForma"] = {k: {"qtd": v["qtd"], "valor": r2(v["valor"])} for k, v in sorted(A["porForma"].items(), key=lambda kv: -kv[1]["valor"])}

    # parcelas por mes de vencimento
    parcelas_por_mes = {}
    for row in linhas.values():
        mk = (row["vencimento"] or "")[:7] or "sem-vencimento"
        parcelas_por_mes.setdefault(mk, []).append([row.get(c) for c in V7_PARCELA_COLS])
    for lst in parcelas_por_mes.values():
        lst.sort(key=lambda x: (x[9] or "", x[19] or ""))

    # v7.1: projecao do mes seguinte = tudo lancado na EVO com vencimento no mes seguinte, sem cancelados
    # (mesma regra do totalValor de receivables_report_next_month*.json)
    nms, nme = next_month_start.isoformat(), next_month_end.isoformat()
    proj_seg = round(sum((row["valor"] or 0) for row in linhas.values()
                         if _v7_norm(row["situacao"]) != "canceled" and row["vencimento"] and nms <= row["vencimento"] <= nme), 2)
    ren_prev = _v71_renovacoes_previstas(key, renovacao, mensalidade)
    ren_ate = _v71_renovacao_ate(renovacao, class_ref or TODAY)
    _v71_log_consultores(key, membros)

    print(f"[{key}] v7: {cur_mk} garantido R$ {G['total']:.2f}, em aberto R$ {A['total']:.2f}, vencido R$ {v_parc:.2f}, recusas do mes R$ {v_rec:.2f}, recuperado R$ {R['total']:.2f}, antecipado R$ {AN['total']:.2f}", file=sys.stderr)
    return {
        "unidade": unit["label"], "key": key, "date": TODAY.isoformat(), "monthKey": cur_mk,
        "posicaoEm": hoje if class_ref else None,
        "projecaoMesSeguinte": {"monthKey": next_month_start.strftime("%Y-%m"),
                                "lancadoNaEvo": proj_seg, "renovacoesPrevistas": ren_prev["total"],
                                "total": round(proj_seg + ren_prev["total"], 2), "renovacoes": ren_prev,
                                "regra": "lancado na EVO com vencimento no mes seguinte (sem cancelados) + renovacoes previstas dos "
                                         "contratos que vencem neste mes e ainda nao renovaram: 70% renovam com +10% sobre o valor atual, "
                                         "30% repostos ao ticket de entrada; a renovacao comeca a pagar no mes seguinte ao vencimento"},
        "renovacao": ren_ate,
        "garantido": G, "emAberto": A, "previsao": r2(G["total"] + A["total"]),
        "cartaoSemAutorizacaoAVencer": dict(CSA, nota="v7.1: TODO cartao com vencimento no mes sem autorizacao (vencido ou nao): fora da previsao e dentro de vencidoRecusado (recusa). Nome do campo mantido do v7."),
        "vencidoRecusado": {
            "parcelasVencidas": {"qtd": len(venc_itens), "total": v_parc,
                                 "regra": "fora cartao: vencido so depois de 3 dias de tolerancia (igual ao v6)"},
            "cartaoSemAutorizacao": {"qtd": len(card_itens), "total": v_card,
                                     "regra": "cartao com vencimento no mes sem autorizacao = recusa (fora da previsao)"},
            "recusasDoMes": {"qtd": len(rec_mes), "total": v_rec},
            "recusasOutrosMeses": {k: {"qtd": v["qtd"], "valor": r2(v["valor"])} for k, v in sorted(rec_outros.items())},
            "totalDoMes": r2(v_parc + v_card + v_rec), "itens": trab,
        },
        "planoEmAberto": {"nota": "anual a parcelar: fora de garantido, em aberto, previsao e vencido",
                          "qtd": len(plano_aberto), "total": r2(sum(x["valor"] for x in plano_aberto)),
                          "itens": sorted(plano_aberto, key=lambda x: x.get("cliente") or "")},
        "saldoDevedorSemClassificacao": {"nota": "saldo devedor sem mensalidade para comparar (nao classificado)",
                                         "qtd": len(saldo_sem), "total": r2(sum(x["valor"] for x in saldo_sem)), "itens": saldo_sem},
        "carteiraDoMesComBaixaNoMes": CM, "recuperado": R, "antecipado": AN,
        "entrouNoMes": r2(CM["total"] + R["total"] + AN["total"]),
        "conferir": {"qtd": len(conferir), "itens": conferir,
                     "regra": "so a MESMA parcela (mesmo id) com duas baixas; nunca por valor ou data parecidos"},
        "_baseBaixa": base_baixa,
        "_parcelasPorMes": parcelas_por_mes, "_baixaRows": baixa_rows,
        "camposAusentes": ausentes,
        "camposEncontrados": dict({f"recebivel.{k}": v for k, v in presenca.items()}, **{f"membro.{k}": v for k, v in pres_membro.items()}),
        "avisos": avisos,
        "contagens": {"recebiveisVencimento": len(raw_venc), "recebiveisBaixa": len(raw_baixa), "parcelasUnicas": len(linhas),
                      "detalhesMembroTelefone": detalhes},
    }


def _v7_read_metas(month_key, arq=None):
    arq = arq or V7_METAS_ARQ  # v7.1: o fechamento do mes anterior le metas_input_prev.json
    out = {"projecao": {"fitness": None, "piscina": None, "total": None}, "fonte": None, "avisos": []}
    try:
        with open(arq) as f:
            m = json.load(f)
    except FileNotFoundError:
        out["avisos"].append(f"{arq} nao encontrado: meta vazia")
        return out
    except Exception as e:
        out["avisos"].append(f"{arq} ilegivel ({e}): meta vazia")
        return out
    if m.get("monthKey") != month_key:
        out["avisos"].append(f"{arq} e de {m.get('monthKey')}, nao de {month_key}: meta vazia")
        return out
    p = m.get("projecao") or {}
    for k in ("fitness", "piscina", "total"):
        out["projecao"][k] = _v7_num(p.get(k))
    out["fonte"] = m.get("fonte")
    out["avisos"] += list(m.get("avisos") or [])
    if all(v is None for v in out["projecao"].values()):
        out["avisos"].append(f"sem projecao de {month_key} na aba Metas: meta vazia")
    return out


def _v7_meta_bloco(projecao, previsao, fonte):
    if projecao is None:
        return {"projecao": None, "fatorEvoBanco": V7_FATOR_EVO_BANCO, "fatorMeta": V7_FATOR_META, "meta": None,
                "previsao": previsao, "faltaParaMeta": None, "passouDaMeta": None, "fonte": fonte}
    meta = round(projecao * V7_FATOR_EVO_BANCO * V7_FATOR_META, 2)
    return {"projecao": projecao, "fatorEvoBanco": V7_FATOR_EVO_BANCO, "fatorMeta": V7_FATOR_META, "meta": meta,
            "previsao": previsao, "faltaParaMeta": round(max(0.0, meta - previsao), 2), "diferenca": round(previsao - meta, 2),
            "passouDaMeta": previsao > meta, "fonte": fonte}


def _v7_comissao(base_baixa, previsao, aliquota, unidade_label):
    total = sum(b["valor"] for b in base_baixa.values())
    por, sem = {}, []
    for b in base_baixa.values():
        c = _v7_consultor_canonico(b.get("consultorEvo"))
        if c:
            p = por.setdefault(c, {"consultor": c, "baseBaixa": 0.0, "qtdClientes": 0})
            p["baseBaixa"] += b["valor"]; p["qtdClientes"] += 1
        else:
            sem.append({"idCliente": b["idCliente"], "cliente": b["cliente"], "unidade": unidade_label,
                        "valor": round(b["valor"], 2), "consultorNaEvo": b.get("consultorEvo")})
    lst = []
    for c in V7_CONSULTORES:
        p = por.get(c, {"consultor": c, "baseBaixa": 0.0, "qtdClientes": 0})
        f = p["baseBaixa"] / total if total else 0.0
        lst.append({"consultor": c, "baseBaixa": round(p["baseBaixa"], 2), "qtdClientes": p["qtdClientes"],
                    "fatia": round(f, 6), "comissao": round(f * previsao * aliquota, 2)})
    sem_v = sum(x["valor"] for x in sem)
    f_sem = sem_v / total if total else 0.0
    sem.sort(key=lambda x: -x["valor"])
    return {"aliquota": aliquota, "baseTotalBaixa": round(total, 2), "porConsultor": lst,
            "semConsultor": {"baseBaixa": round(sem_v, 2), "fatia": round(f_sem, 6), "qtdClientes": len(sem),
                             "comissaoARedistribuir": round(f_sem * previsao * aliquota, 2),
                             "nota": "NAO fica com a empresa: lista para redistribuir", "clientes": sem}}


def finalize_extras(TODAY_ref=None):
    """Grava os arquivos novos do v7 com o que cada unidade calculou. Chamado no fim do script (finally)."""
    if os.environ.get("COLETA_DATA_REF"):
        return
    manifest = []
    try:
        if not EXTRAS_OUT:
            print(f"[v7] nenhum extra calculado ({EXTRAS_ERR or 'sem unidades'})", file=sys.stderr)
            return
        any_u = next(iter(EXTRAS_OUT.values()))
        mk, hoje = any_u["monthKey"], any_u["date"]
        metas = _v7_read_metas(mk)
        proj = dict(metas["projecao"])
        cong = _v71_projecao_congelada(mk, hoje)
        if cong:
            # v7.1 (item 2): no dia 1o a projecao congelada hoje substitui a da aba Metas (a rotina grava a mesma
            # projecao na aba "Metas do mes"; a partir do dia 2 o passo 2c ja le de la)
            for k, pms in cong.items():
                proj[k] = pms["total"]
            proj["total"] = None
            metas["fonte"] = "Projeção congelada na coleta do dia 1º (lançado na EVO + renovações previstas)"
            metas["avisos"] = [a for a in metas["avisos"] if "sem projecao" not in a and "meta vazia" not in a]
        prev_total = round(sum(u["previsao"] for u in EXTRAS_OUT.values()), 2)
        if proj["total"] is None and proj["fitness"] is not None and proj["piscina"] is not None:
            proj["total"] = round(proj["fitness"] + proj["piscina"], 2)
        meta_cons = _v7_meta_bloco(proj["total"], prev_total, metas["fonte"])
        if len(EXTRAS_OUT) < 2:
            meta_cons["aviso"] = "so uma unidade foi calculada hoje: previsao consolidada incompleta"
        cons_por = {}
        cons_sem = []
        cons_sem_com = 0.0
        for ukey, u in EXTRAS_OUT.items():
            suf = "" if ukey == "fitness" else "_piscina"
            col_sfx = "" if ukey == "fitness" else "Piscina"
            meta_u = _v7_meta_bloco(proj.get(ukey), u["previsao"], metas["fonte"])
            if meta_u["meta"] is not None:
                aliq, por_q = (V7_ALIQ_META if meta_u["passouDaMeta"] else V7_ALIQ_BASE), "meta da unidade"
            elif meta_cons["meta"] is not None and len(EXTRAS_OUT) == 2:
                aliq, por_q = (V7_ALIQ_META if meta_cons["passouDaMeta"] else V7_ALIQ_BASE), "meta consolidada (Academia + Piscina)"
            else:
                aliq, por_q = V7_ALIQ_BASE, "sem meta: aliquota base provisoria"
            com = _v7_comissao(u["_baseBaixa"], u["previsao"], aliq, u["unidade"])
            com["aliquotaDecididaPor"] = por_q
            com["nota"] = "previa diaria: fatia = baixa no mes (garantido + recuperado + antecipado) dos clientes do consultor / total com baixa no mes; comissao = fatia x previsao x aliquota. No fechamento a comissao oficial continua pelo banco."
            for p in com["porConsultor"]:
                c = cons_por.setdefault(p["consultor"], {"consultor": p["consultor"], "baseBaixa": 0.0, "qtdClientes": 0, "comissao": 0.0})
                c["baseBaixa"] += p["baseBaixa"]; c["qtdClientes"] += p["qtdClientes"]; c["comissao"] += p["comissao"]
            cons_sem += com["semConsultor"]["clientes"]
            cons_sem_com += com["semConsultor"]["comissaoARedistribuir"]

            # parcelas e baixas (documentos divididos por tamanho)
            meta_doc = {"unidade": u["unidade"], "date": hoje, "updatedAt": datetime.datetime.utcnow().isoformat() + "Z"}
            docs_parc = {}
            for vmk, rows in sorted(u["_parcelasPorMes"].items()):
                if ukey in EXTRAS_PREV_OUT and vmk == EXTRAS_PREV_OUT[ukey]["monthKey"]:
                    continue  # v7.1: dias 1 a 5, o mes anterior e gravado pelo fechamento (posicao final)
                partes = _v7_split_rows(dict(meta_doc, monthKey=vmk, tipo="parcelas por mes de vencimento"), vmk, V7_PARCELA_COLS, rows)
                docs_parc[vmk] = len(partes)
                for p in partes:
                    fn = f"parcelas_evo{suf}_{vmk}_p{p['parte']}.json"
                    with open(fn, "w") as f:
                        json.dump(p, f, ensure_ascii=False)
                    manifest.append([fn, "parcelasEvo" + col_sfx, p["docId"]])
            bcols = ["idParcela", "idCliente", "cliente", "consultorEvo", "descricao", "vencimento", "dataBaixa", "valor",
                     "valorPago", "forma", "situacao", "origem", "idContrato", "idVenda"]
            partes_b = _v7_split_rows(dict(meta_doc, monthKey=mk, tipo="baixas no mes (qualquer vencimento)"), mk, bcols, u["_baixaRows"])
            for p in partes_b:
                fn = f"baixas_evo{suf}_p{p['parte']}.json"
                with open(fn, "w") as f:
                    json.dump(p, f, ensure_ascii=False)
                manifest.append([fn, "baixasEvo" + col_sfx, p["docId"]])

            resumo = {k: v for k, v in u.items() if not k.startswith("_")}
            resumo.update({"docId": mk, "meta": meta_u, "metaConsolidada": meta_cons, "comissaoParcial": com,
                           "documentosParcelas": docs_parc, "documentosBaixas": len(partes_b),
                           "avisos": u["avisos"] + metas["avisos"],
                           "regras": V71_REGRAS,
                           "updatedAt": datetime.datetime.utcnow().isoformat() + "Z"})
            fn = f"resumo_mes{suf}.json"
            if len(json.dumps(resumo, ensure_ascii=False).encode()) > V7_MAX_DOC_BYTES:
                listas = {"docId": f"{mk}-listas", "monthKey": mk, "date": hoje, "unidade": u["unidade"],
                          "vencidoRecusadoItens": resumo["vencidoRecusado"].pop("itens"),
                          "semConsultorClientes": resumo["comissaoParcial"]["semConsultor"].pop("clientes"),
                          "planoEmAbertoItens": resumo["planoEmAberto"].pop("itens"),
                          "updatedAt": resumo["updatedAt"]}
                resumo["listasEmDocumentoSeparado"] = f"{mk}-listas"
                fl = f"resumo_mes_listas{suf}.json"
                with open(fl, "w") as f:
                    json.dump(listas, f, ensure_ascii=False)
                manifest.append([fl, "resumoMes" + col_sfx, listas["docId"]])
            with open(fn, "w") as f:
                json.dump(resumo, f, ensure_ascii=False)
            manifest.append([fn, "resumoMes" + col_sfx, mk])

            dia = {"docId": hoje, "date": hoje, "monthKey": mk, "unidade": u["unidade"],
                   "garantido": u["garantido"]["total"], "emAberto": u["emAberto"]["total"], "previsao": u["previsao"],
                   "vencidoRecusado": u["vencidoRecusado"]["totalDoMes"], "recuperado": u["recuperado"]["total"],
                   "antecipado": u["antecipado"]["total"], "entrouNoMes": u["entrouNoMes"],
                   "meta": meta_u["meta"], "faltaParaMeta": meta_u["faltaParaMeta"], "aliquota": aliq,
                   "comissaoParcial": {p["consultor"]: p["comissao"] for p in com["porConsultor"]},
                   "semConsultorARedistribuir": com["semConsultor"]["comissaoARedistribuir"],
                   "projecaoMesSeguinte": _v71_proj_dia(u["projecaoMesSeguinte"]),
                   "updatedAt": datetime.datetime.utcnow().isoformat() + "Z"}
            fd = f"resumo_dia{suf}.json"
            with open(fd, "w") as f:
                json.dump(dia, f, ensure_ascii=False)
            manifest.append([fd, "resumoDiario" + col_sfx, hoje])
            print(f"[{ukey}] v7: previsao R$ {u['previsao']:.2f}, meta {meta_u['meta']}, aliquota {aliq} ({por_q}), comissao {[(p['consultor'], p['comissao']) for p in com['porConsultor']]}, sem consultor {com['semConsultor']['qtdClientes']} clientes", file=sys.stderr)

        ausentes = sorted({a for u in EXTRAS_OUT.values() for a in u["camposAusentes"]})
        print(f"[v7] campos ausentes na EVO: {', '.join(ausentes) if ausentes else 'nenhum'}", file=sys.stderr)
        for a in metas["avisos"]:
            print(f"[v7] aviso: {a}", file=sys.stderr)
        if len(EXTRAS_OUT) < 2:
            # com uma unidade so, o consolidado ficaria incompleto: nao grava (o do dia anterior continua valendo)
            print(f"[v7] consolidado NAO gerado: so {', '.join(EXTRAS_OUT)} foi calculada", file=sys.stderr)
            return
        cons = {"docId": mk, "monthKey": mk, "date": hoje, "unidades": sorted(u["unidade"] for u in EXTRAS_OUT.values()),
                "unidadesComErro": EXTRAS_ERR,
                "garantido": round(sum(u["garantido"]["total"] for u in EXTRAS_OUT.values()), 2),
                "emAberto": round(sum(u["emAberto"]["total"] for u in EXTRAS_OUT.values()), 2),
                "previsao": prev_total,
                "vencidoRecusado": round(sum(u["vencidoRecusado"]["totalDoMes"] for u in EXTRAS_OUT.values()), 2),
                "planoEmAberto": round(sum(u["planoEmAberto"]["total"] for u in EXTRAS_OUT.values()), 2),
                "recuperado": round(sum(u["recuperado"]["total"] for u in EXTRAS_OUT.values()), 2),
                "antecipado": round(sum(u["antecipado"]["total"] for u in EXTRAS_OUT.values()), 2),
                "entrouNoMes": round(sum(u["entrouNoMes"] for u in EXTRAS_OUT.values()), 2),
                "projecaoMesSeguinte": _v71_soma_proj(EXTRAS_OUT.values()),
                "meta": meta_cons,
                "comissaoParcial": {"porConsultor": [dict(v, baseBaixa=round(v["baseBaixa"], 2), comissao=round(v["comissao"], 2)) for v in cons_por.values()],
                                    "semConsultor": {"qtdClientes": len(cons_sem), "baseBaixa": round(sum(x["valor"] for x in cons_sem), 2),
                                                     "comissaoARedistribuir": round(cons_sem_com, 2)},
                                    "nota": "soma das unidades; previa diaria — no fechamento a comissao oficial continua pelo banco"},
                "camposAusentes": ausentes, "avisos": metas["avisos"],
                "updatedAt": datetime.datetime.utcnow().isoformat() + "Z"}
        with open("resumo_mes_consolidado.json", "w") as f:
            json.dump(cons, f, ensure_ascii=False)
        manifest.append(["resumo_mes_consolidado.json", "resumoMesConsolidado", mk])
        dia_c = {"docId": hoje, "date": hoje, "monthKey": mk, "garantido": cons["garantido"], "emAberto": cons["emAberto"],
                 "previsao": cons["previsao"], "vencidoRecusado": cons["vencidoRecusado"], "recuperado": cons["recuperado"],
                 "antecipado": cons["antecipado"], "entrouNoMes": cons["entrouNoMes"], "meta": meta_cons["meta"],
                 "faltaParaMeta": meta_cons["faltaParaMeta"],
                 "comissaoParcial": {p["consultor"]: p["comissao"] for p in cons["comissaoParcial"]["porConsultor"]},
                 "projecaoMesSeguinte": cons["projecaoMesSeguinte"],
                 "unidades": cons["unidades"], "updatedAt": cons["updatedAt"]}
        with open("resumo_dia_consolidado.json", "w") as f:
            json.dump(dia_c, f, ensure_ascii=False)
        manifest.append(["resumo_dia_consolidado.json", "resumoDiarioConsolidado", hoje])
    except Exception as e:
        print(f"[v7] extras falharam ao gravar (arquivos do v6 nao foram afetados): {e!r}", file=sys.stderr)
        manifest = []
    finally:
        try:
            _v71_extras(manifest)
        except Exception as e:
            print(f"[v7.1] extras da v7.1 falharam (v6 e v7 nao foram afetados): {e!r}", file=sys.stderr)
        try:
            with open("v7_manifest.json", "w") as f:
                json.dump(manifest, f, ensure_ascii=False)
        except Exception:
            pass
        if manifest:
            print(f"[v7] extras done. {len(manifest)} arquivos novos ({', '.join(sorted(EXTRAS_OUT))}){' | falharam: ' + ', '.join(sorted(EXTRAS_ERR)) if EXTRAS_ERR else ''}", file=sys.stderr)


# =====================================================================================================
# v7.1 — fechamento do mes anterior (A) e linhas da aba "Mes atual" (B). So acrescenta.
# =====================================================================================================
V71_METAS_PREV_ARQ = "metas_input_prev.json"
V71_MES_ATUAL_ARQ = "mes_atual_linhas.json"
V71_COL_CONSULTOR = "Comissão parcial – {}"
V71_MES_ATUAL_COLS = ["Data", "Unidade", "Garantido", "Em aberto", "Previsão", "Vencido/recusado", "Meta",
                      "Comissão parcial – Raquel Ribeiro", "Comissão parcial – sem consultor", "Projeção do mês seguinte", "Fonte"]
V71_REGRAS = {
    "garantido": "vencimento no mes e ja cobrado: cartao autorizado (vencimento passado ou futuro dentro do mes) + demais formas com baixa",
    "emAberto": "vencimento no mes, sem baixa, forma diferente de cartao (boleto, Pix...), ainda dentro de vencimento + 3 dias",
    "previsao": "garantido + em aberto (comparar com a meta)",
    "vencidoRecusado": "parcelas do mes vencidas sem baixa (fora cartao, depois de 3 dias) + cartao do mes sem autorizacao + saldo devedor de cobranca recusada (ligado ao mes do dinheiro = cobranca + 30 dias)",
    "projecaoMesSeguinte": "lancado na EVO com vencimento no mes seguinte (sem cancelados) + renovacoes previstas dos contratos que vencem no mes e nao renovaram (70% x valor atual x 1,10 + 30% x ticket de entrada)",
    "renovacaoAteHoje": "renovados / contratos que ja venceram ate a data",
    "saldoDevedor": "ate 1,5 mensalidade = cobranca recusada; acima = plano em aberto (anual a parcelar), salvo se a EVO informar a origem",
    "recuperado": "baixa no mes, vencimento anterior", "antecipado": "baixa no mes, vencimento posterior",
    "meta": "projecao do mes x fator EVO x banco x 1,005",
    "posicaoFinal": "fechamento: refeito nos dias 1 a 5 do mes seguinte, com os pagamentos lancados ate hoje; vencido/a vencer pela posicao do dia 1o (boleto/Pix dos 3 ultimos dias ainda na tolerancia contam como em aberto); cartao sem autorizacao = recusa",
    "cartaoSemAutorizacao": "todo cartao com vencimento no mes sem autorizacao = recusa: fora da previsao, dentro de vencidoRecusado e da lista de trabalho",
    "vencidoTolerancia": "fora cartao, vencido so 3 dias depois do vencimento (igual ao v6)",
}


def _v71_soma_proj(us):
    us = list(us)
    lanc = round(sum(u["projecaoMesSeguinte"]["lancadoNaEvo"] for u in us), 2)
    ren = round(sum(u["projecaoMesSeguinte"]["renovacoesPrevistas"] for u in us), 2)
    return {"monthKey": us[0]["projecaoMesSeguinte"]["monthKey"], "lancadoNaEvo": lanc, "renovacoesPrevistas": ren,
            "total": round(lanc + ren, 2)}


def _v71_proj_dia(pms):
    return {k: pms[k] for k in ("monthKey", "lancadoNaEvo", "renovacoesPrevistas", "total")}


def _v71_projecao_congelada(mk, hoje_iso):
    """Dia 1o: projecao do mes que comeca (mk) = projecaoMesSeguinte do fechamento do mes anterior."""
    if not hoje_iso or hoje_iso[8:10] != "01" or os.environ.get("COLETA_DATA_REF"):
        return {}
    out = {}
    for k, u in EXTRAS_PREV_OUT.items():
        pms = u.get("projecaoMesSeguinte") or {}
        if pms.get("monthKey") == mk and pms.get("total") is not None:
            out[k] = pms
    return out


def _v71_dump(fn, obj):
    with open(fn, "w") as f:
        json.dump(obj, f, ensure_ascii=False)


def _v71_finalize_prev(manifest):
    if EXTRAS_PREV_ERR:
        for k, e in EXTRAS_PREV_ERR.items():
            print(f"[v7.1] fechamento do mes anterior falhou em {k}: {e}", file=sys.stderr)
    if not EXTRAS_PREV_OUT:
        return
    n0 = len(manifest)
    any_u = next(iter(EXTRAS_PREV_OUT.values()))
    mk, hoje = any_u["monthKey"], any_u["date"]
    agora = datetime.datetime.utcnow().isoformat() + "Z"
    metas = _v7_read_metas(mk, V71_METAS_PREV_ARQ)
    proj = dict(metas["projecao"])
    prev_total = round(sum(u["previsao"] for u in EXTRAS_PREV_OUT.values()), 2)
    if proj["total"] is None and proj["fitness"] is not None and proj["piscina"] is not None:
        proj["total"] = round(proj["fitness"] + proj["piscina"], 2)
    meta_cons = _v7_meta_bloco(proj["total"], prev_total, metas["fonte"])
    pos = {"posicao": "final", "fechamento": True, "posicaoEm": any_u.get("posicaoEm"), "geradoEm": hoje}
    cons_por, cons_sem, cons_sem_com = {}, [], 0.0
    for ukey, u in EXTRAS_PREV_OUT.items():
        suf = "" if ukey == "fitness" else "_piscina"
        col_sfx = "" if ukey == "fitness" else "Piscina"
        meta_u = _v7_meta_bloco(proj.get(ukey), u["previsao"], metas["fonte"])
        if meta_u["meta"] is not None:
            aliq, por_q = (V7_ALIQ_META if meta_u["passouDaMeta"] else V7_ALIQ_BASE), "meta da unidade"
        elif meta_cons["meta"] is not None and len(EXTRAS_PREV_OUT) == 2:
            aliq, por_q = (V7_ALIQ_META if meta_cons["passouDaMeta"] else V7_ALIQ_BASE), "meta consolidada (Academia + Piscina)"
        else:
            aliq, por_q = V7_ALIQ_BASE, "sem meta: aliquota base provisoria"
        com = _v7_comissao(u["_baseBaixa"], u["previsao"], aliq, u["unidade"])
        com["aliquotaDecididaPor"] = por_q
        com["nota"] = "posicao final do mes pela EVO: fatia = baixa no mes dos clientes do consultor / total com baixa no mes; comissao = fatia x previsao x aliquota. A comissao oficial continua pelo banco."
        for p in com["porConsultor"]:
            c = cons_por.setdefault(p["consultor"], {"consultor": p["consultor"], "baseBaixa": 0.0, "qtdClientes": 0, "comissao": 0.0})
            c["baseBaixa"] += p["baseBaixa"]; c["qtdClientes"] += p["qtdClientes"]; c["comissao"] += p["comissao"]
        cons_sem += com["semConsultor"]["clientes"]
        cons_sem_com += com["semConsultor"]["comissaoARedistribuir"]

        meta_doc = dict({"unidade": u["unidade"], "date": hoje, "updatedAt": agora}, **pos)
        # parcelas: so as com vencimento no proprio mes anterior (doc_id AAAA-MM-pN)
        partes_p = _v7_split_rows(dict(meta_doc, monthKey=mk, tipo="parcelas por mes de vencimento"), mk, V7_PARCELA_COLS,
                                  u["_parcelasPorMes"].get(mk, []))
        for p in partes_p:
            fn = f"parcelas_evo_prev{suf}_{mk}_p{p['parte']}.json"
            _v71_dump(fn, p)
            manifest.append([fn, "parcelasEvo" + col_sfx, p["docId"]])
        bcols = ["idParcela", "idCliente", "cliente", "consultorEvo", "descricao", "vencimento", "dataBaixa", "valor",
                 "valorPago", "forma", "situacao", "origem", "idContrato", "idVenda"]
        partes_b = _v7_split_rows(dict(meta_doc, monthKey=mk, tipo="baixas no mes (qualquer vencimento)"), mk, bcols, u["_baixaRows"])
        for p in partes_b:
            fn = f"baixas_evo_prev{suf}_p{p['parte']}.json"
            _v71_dump(fn, p)
            manifest.append([fn, "baixasEvo" + col_sfx, p["docId"]])

        resumo = {k: v for k, v in u.items() if not k.startswith("_")}
        resumo.update(pos)
        resumo.update({"docId": mk, "meta": meta_u, "metaConsolidada": meta_cons, "comissaoParcial": com,
                       "documentosParcelas": {mk: len(partes_p)}, "documentosBaixas": len(partes_b),
                       "avisos": u["avisos"] + metas["avisos"], "regras": V71_REGRAS, "updatedAt": agora})
        if len(json.dumps(resumo, ensure_ascii=False).encode()) > V7_MAX_DOC_BYTES:
            listas = {"docId": f"{mk}-listas", "monthKey": mk, "date": hoje, "unidade": u["unidade"],
                      "vencidoRecusadoItens": resumo["vencidoRecusado"].pop("itens"),
                      "semConsultorClientes": resumo["comissaoParcial"]["semConsultor"].pop("clientes"),
                      "planoEmAbertoItens": resumo["planoEmAberto"].pop("itens"), "updatedAt": agora}
            listas.update(pos)
            resumo["listasEmDocumentoSeparado"] = f"{mk}-listas"
            fl = f"resumo_mes_prev_listas{suf}.json"
            _v71_dump(fl, listas)
            manifest.append([fl, "resumoMes" + col_sfx, listas["docId"]])
        fn = f"resumo_mes_prev{suf}.json"
        _v71_dump(fn, resumo)
        manifest.append([fn, "resumoMes" + col_sfx, mk])
        print(f"[{ukey}] v7.1 fechamento {mk}: previsao R$ {u['previsao']:.2f}, meta {meta_u['meta']}, aliquota {aliq} ({por_q}), comissao {[(p['consultor'], p['comissao']) for p in com['porConsultor']]}", file=sys.stderr)

    for a in metas["avisos"]:
        print(f"[v7] aviso: fechamento {mk}: {a}", file=sys.stderr)
    if len(EXTRAS_PREV_OUT) == 2:
        U = EXTRAS_PREV_OUT.values()
        cons = dict({"docId": mk, "monthKey": mk, "date": hoje, "unidades": sorted(u["unidade"] for u in U),
                     "unidadesComErro": EXTRAS_PREV_ERR,
                     "garantido": round(sum(u["garantido"]["total"] for u in U), 2),
                     "emAberto": round(sum(u["emAberto"]["total"] for u in U), 2),
                     "previsao": prev_total,
                     "vencidoRecusado": round(sum(u["vencidoRecusado"]["totalDoMes"] for u in U), 2),
                     "planoEmAberto": round(sum(u["planoEmAberto"]["total"] for u in U), 2),
                     "recuperado": round(sum(u["recuperado"]["total"] for u in U), 2),
                     "antecipado": round(sum(u["antecipado"]["total"] for u in U), 2),
                     "entrouNoMes": round(sum(u["entrouNoMes"] for u in U), 2),
                     "projecaoMesSeguinte": _v71_soma_proj(U),
                     "meta": meta_cons,
                     "comissaoParcial": {"porConsultor": [dict(v, baseBaixa=round(v["baseBaixa"], 2), comissao=round(v["comissao"], 2)) for v in cons_por.values()],
                                         "semConsultor": {"qtdClientes": len(cons_sem), "baseBaixa": round(sum(x["valor"] for x in cons_sem), 2),
                                                          "comissaoARedistribuir": round(cons_sem_com, 2)},
                                         "nota": "soma das unidades; posicao final pela EVO — a comissao oficial continua pelo banco"},
                     "camposAusentes": sorted({a for u in U for a in u["camposAusentes"]}), "avisos": metas["avisos"],
                     "updatedAt": agora}, **pos)
        _v71_dump("resumo_mes_prev_consolidado.json", cons)
        manifest.append(["resumo_mes_prev_consolidado.json", "resumoMesConsolidado", mk])
    else:
        print(f"[v7.1] fechamento {mk}: consolidado NAO gerado: so {', '.join(EXTRAS_PREV_OUT)} foi calculada", file=sys.stderr)
    print(f"[v7.1] fechamento {mk} done. {len(manifest) - n0} arquivos ({', '.join(sorted(EXTRAS_PREV_OUT))})", file=sys.stderr)


def _v71_mes_atual(manifest):
    """Linhas da aba "Mes atual" (Base de Numeros) com os numeros do resumo do mes gravado hoje."""
    arqs = {x[0] for x in manifest}
    cols = list(V71_MES_ATUAL_COLS) + [V71_COL_CONSULTOR.format(c) for c in V7_CONSULTORES
                                       if V71_COL_CONSULTOR.format(c) not in V71_MES_ATUAL_COLS]
    fonte = "EVO (coleta diária do Dashboard 4Live, v7.1)"

    def linha(data, unidade, r, meta, com_por, sem_com):
        d = {"Data": data, "Unidade": unidade, "Garantido": r["garantido"], "Em aberto": r["emAberto"],
             "Previsão": r["previsao"], "Vencido/recusado": r["vencidoRecusado"], "Meta": meta,
             "Comissão parcial – sem consultor": sem_com, "Projeção do mês seguinte": r["projecaoMesSeguinte"]["total"],
             "Fonte": fonte}
        for c in V7_CONSULTORES:
            d[V71_COL_CONSULTOR.format(c)] = com_por.get(c, 0.0)
        return d

    linhas, mk = [], None
    for suf, lbl in (("", "Fitness"), ("_piscina", "Piscina")):
        fn = f"resumo_mes{suf}.json"
        if fn not in arqs:
            continue
        r = json.load(open(fn))
        mk = r["monthKey"]
        flat = dict(r, garantido=r["garantido"]["total"], emAberto=r["emAberto"]["total"],
                    vencidoRecusado=r["vencidoRecusado"]["totalDoMes"])
        cp = r["comissaoParcial"]
        linhas.append(linha(r["date"], lbl, flat, r["meta"]["meta"], {p["consultor"]: p["comissao"] for p in cp["porConsultor"]},
                            cp["semConsultor"]["comissaoARedistribuir"]))
    if "resumo_mes_consolidado.json" in arqs and len(linhas) == 2:
        r = json.load(open("resumo_mes_consolidado.json"))
        cp = r["comissaoParcial"]
        linhas.append(linha(r["date"], "Total", r, r["meta"]["meta"], {p["consultor"]: p["comissao"] for p in cp["porConsultor"]},
                            cp["semConsultor"]["comissaoARedistribuir"]))
    if not linhas:
        print("[v7.1] mes atual: nenhuma linha (resumo do mes nao foi gerado hoje)", file=sys.stderr)
        return
    _v71_dump(V71_MES_ATUAL_ARQ, {"monthKey": mk, "colunas": cols, "linhas": linhas})
    print(f"[v7.1] mes atual: {len(linhas)} linhas ({', '.join(l['Unidade'] for l in linhas)})", file=sys.stderr)


def _v71_meta_congelada():
    if not EXTRAS_PREV_OUT:
        return
    hoje = next(iter(EXTRAS_PREV_OUT.values()))["date"]
    mk = hoje[:7]
    cong = _v71_projecao_congelada(mk, hoje)
    if not cong:
        return
    lbl = {"fitness": "Fitness", "piscina": "Piscina"}
    doc = {"monthKey": mk, "geradoEm": hoje,
           "regra": "projecao do mes que comeca = lancado na EVO + renovacoes previstas (70% +10%, 30% ticket de entrada); meta = projecao x 1,005",
           "unidades": [{"unidade": lbl[k], "valor": p["total"], "lancadoNaEvo": p["lancadoNaEvo"],
                         "renovacoesPrevistas": p["renovacoesPrevistas"]} for k, p in sorted(cong.items())]}
    _v71_dump("meta_congelada.json", doc)
    print(f"[v7.1] meta congelada {mk}: " + ", ".join(f"{x['unidade']} R$ {x['valor']:.2f}" for x in doc["unidades"]), file=sys.stderr)


def _v71_extras(manifest):
    try:
        _v71_meta_congelada()
    except Exception as e:
        print(f"[v7.1] meta congelada falhou: {e!r}", file=sys.stderr)
    try:
        _v71_finalize_prev(manifest)
    except Exception as e:
        print(f"[v7.1] fechamento do mes anterior falhou ao gravar (mes corrente nao foi afetado): {e!r}", file=sys.stderr)
    try:
        _v71_mes_atual(manifest)
    except Exception as e:
        print(f"[v7.1] mes atual falhou (nada mais foi afetado): {e!r}", file=sys.stderr)




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

    # Pedido do usuario em 2026-10-01: modo RECALCULO de mes fechado. Com COLETA_DATA_REF=AAAA-MM-DD
    # (ultimo dia do mes), o script calcula o relatorio de clientes daquele mes com as regras atuais
    # (cancelamentos sem trocas, valor perdido = saldo, renovacao com prazo de 30 dias, leads) e grava
    # client_report_backfill*.json, sales_report_backfill*.json e receivables_report_backfill*.json —
    # nada de snapshot nem listas (vendas e recebiveis incluidos em 2026-11-09). Usado para
    # corrigir julho e agosto/2026, gravados em 16/09 com as regras antigas.
    TODAY = datetime.date.fromisoformat(os.environ["COLETA_DATA_REF"]) if os.environ.get("COLETA_DATA_REF") else datetime.date.today()
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
    _cut7 = (TODAY - datetime.timedelta(days=6)).isoformat()

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
    # Pedido do usuario em 2026-10-08: upgrade/downgrade de plano nao e contrato vencendo. O EVO marca o
    # contrato NOVO com membershipSwapData.flMembershipSwapped = True e aponta o antigo em
    # idMemberMembershipSource; o antigo (encerrado na data da troca) sai de "Contratos vencendo"
    # (igual ao relatorio de renovacao do EVO — ex. Piscina out/2026: Hugo, Manoel e Marcelo).
    swapped_sources = {(c.get("membershipSwapData") or {}).get("idMemberMembershipSource") for c in all_contracts
                       if (c.get("membershipSwapData") or {}).get("flMembershipSwapped")}
    swapped_sources.discard(None)
    # Pedido do usuario em 2026-10-12: a VENDA do contrato novo de um upgrade/downgrade (o mesmo
    # flMembershipSwapped acima) tambem e troca de contrato no relatorio de vendas, nao Renovacao/Novo.
    # O heuristico de build_contract_swap_report nao pegava esses casos porque o EVO nao cancela o
    # contrato antigo nem cobra adesao (Piscina out/2026: Hugo 12591, Manoel 12592, Marcelo 12598).
    swapped_sale_ids = {c.get("idSale") for c in all_contracts
                        if (c.get("membershipSwapData") or {}).get("flMembershipSwapped")}
    swapped_sale_ids.discard(None)

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
    # Pedido do usuario em 2026-09-30: "Valor perdido" desconta as trocas internas de plano (igual ao
    # numero de Cancelamentos, que ja desconta). O valor com as trocas fica em cancelledValueBruto.
    cancelled_value_bruto = sum(parse_value_br(r.get("Valor")) for r in cancelled)
    cancelled_value = sum(parse_value_br(r.get("Valor")) for r in cancelled
                          if categorize_reason(r.get("MotivoCancelamento")) != "Troca interna de plano")
    reasons = {}
    for r in cancelled:
        cat = categorize_reason(r.get("MotivoCancelamento"))
        reasons[cat] = reasons.get(cat, 0) + 1
    internal_plan_changes = reasons.get("Troca interna de plano", 0)

    total_leads = len(prospects)
    converted = [r for r in prospects if str(r.get("Status")).strip().upper() == "CLIENTE"]

    turma_cache = {}

    def build_renewal_report(range_start, range_end):
        due = []
        for idm, contracts in contracts_by_member.items():
            for c in contracts:
                # Pedido do usuario em 2026-10-08: o EVO passou a marcar o contrato que terminou como
                # status 3 (encerrado) — antes so o 1 (ativo) entrava, e os contratos que ja venceram
                # sumiam da lista (setembro zerado; outubro sem os que venceram de 01 a 07/10).
                # Entram 1 (ativo) e 3 (encerrado); 2 (cancelado) continua fora.
                if c["_end"] is None or c.get("statusMemberMembership") not in (1, 3):
                    continue
                if not (range_start <= c["_end"] <= range_end):
                    continue
                if is_transfer_record(c):
                    continue
                if c.get("idMemberMemberShip") in swapped_sources:
                    continue
                if is_excluded_from_renewals(c.get("nameMembership")):
                    continue
                # Pedido do usuario em 2026-09-29: no FITNESS, troca de plano na renovacao (ex. Fitness
                # Anual DCC -> Fitness Assinatura Plus, 1 dia depois) E renovacao — o contrato antigo
                # continua em "Contratos vencendo" e conta como renovado. Validado contra o relatorio
                # de renovacao do EVO de setembro/2026 (34 contratos, 19 renovados; 17 renovacoes com
                # troca de plano estavam sumindo). A regra de troca de modalidade continua so na
                # Piscina, onde foi validada com o usuario (Joao Vicente, Lunna, Leonardo).
                # Pedido do usuario em 2026-09-30: a Piscina passa a usar a MESMA regra do Fitness
                # (troca de plano na renovacao = renovado). A regra is_internal_plan_swap tirava da
                # lista 18 contratos em setembro/2026, a maioria renovacao real (migracao Anual DCC ->
                # Assinatura Plus). Validado contra o relatorio de renovacao do EVO de setembro/2026 da
                # Piscina (26 contratos, 20 renovados, 6 nao renovados). So Lunna e Joao Vicente ficam
                # fora, por exclusao manual (_MANUAL_RENEWAL_EXCLUSIONS).
                if is_manually_excluded_from_renewals(idm, c):
                    continue
                # Pedido do usuario em 2026-09-29: renovado = existe QUALQUER contrato novo comecando
                # depois do fim deste, mesmo com outro plano. Excecao: "Cortesia 15 dias" nao conta
                # como renovacao — se o contrato seguinte for cortesia, olha-se o contrato pago que
                # vem depois dela (se houver). A propria cortesia continua fora da lista (keyword
                # "cortesia" em _RENEWAL_EXCLUDE_KEYWORDS), entao o que vale e o plano anterior a ela.
                # Pedido do usuario em 2026-10-01: so conta como renovacao se o contrato novo comecar ate
                # RENOVACAO_PRAZO_DIAS (30) dias depois do fim deste — depois disso e Retorno, igual na
                # classificacao das vendas. Renovacao feita no mes seguinte ao vencimento (dentro dos 30
                # dias) conta para o mes do vencimento e fica marcada como "renovadoTardio".
                _limite = c["_end"] + datetime.timedelta(days=RENOVACAO_PRAZO_DIAS)
                later = [x for x in contracts if x["_start"] and c["_end"] < x["_start"] <= _limite and not is_transfer_record(x)
                         and "cortesia" not in _strip_accents((x.get("nameMembership") or "").strip().lower())]
                _ren_start = min(x["_start"] for x in later) if later else None
                _tardio = bool(_ren_start and (_ren_start.year, _ren_start.month) > (c["_end"].year, c["_end"].month))
                due.append((idm, c, len(later) > 0, _tardio, _ren_start))

        ids_needed = sorted({d[0] for d in due if d[0] not in member_cache})
        if ids_needed:
            chunk = ",".join(str(i) for i in ids_needed)
            extra_members = json.loads(api_get_raw(f"{API_BASE_V2}/members?idsMembers={chunk}&take=1000"))
            for m in extra_members:
                member_cache[m["idMember"]] = {"consultor": m.get("nameEmployeeConsultant"), "professor": m.get("nameEmployeeInstructor")}

        # Pedido do usuario em 2026-10-22: na Piscina o EVO nao traz professor no cadastro, entao o
        # professor da renovacao vem da TURMA que o aluno frequentou (chamada das aulas nos ultimos
        # 120 dias; vale a turma mais recente — aulas ate 14 dias antes da ultima). Fica gravado no
        # item (professorTurma/turma), assim quem ja saiu da turma continua com o professor no mes.
        def turma_do_aluno(idm):
            if idm in turma_cache:
                return turma_cache[idm]
            res = (None, None)
            try:
                recs = get_member_activity_sessions(api_get_raw, idm, TODAY - datetime.timedelta(days=119), TODAY)
                recs = [r for r in recs if (r.get("date") or "")[:10]]
                if recs:
                    _ult = max(r["date"][:10] for r in recs)
                    _corte = (datetime.date.fromisoformat(_ult) - datetime.timedelta(days=13)).isoformat()
                    rec = [r for r in recs if r["date"][:10] >= _corte]
                    res = (", ".join(sorted(set((r.get("instructor") or "").strip() for r in rec if (r.get("instructor") or "").strip()))) or None,
                           ", ".join(sorted(set(r.get("activitieName") for r in rec if r.get("activitieName")))) or None)
            except Exception as ex:
                print(f"[{unit['key']}] turma do aluno {idm} falhou: {ex}", file=sys.stderr)
            turma_cache[idm] = res
            return res

        items = []
        renewed = not_renewed_n = 0
        for idm, c, is_renewed, is_tardio, ren_start in due:
            e = member_cache.get(idm, {})
            prof_turma, turma_nome = turma_do_aluno(idm) if unit["key"] == "piscina" else (None, None)
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
                # Pedido do usuario em 2026-10-01: renovou depois do mes do vencimento (dentro dos 30 dias).
                "renovadoTardio": is_tardio, "renovadoEm": ren_start.isoformat() if ren_start else None,
                "professorTurma": prof_turma, "turma": turma_nome,
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
    _prev_m_end = month_start - datetime.timedelta(days=1)
    if (TODAY - _prev_m_end).days <= RENOVACAO_PRAZO_DIAS + 1:
        CLIENT_PREV_OUT[unit["key"]] = {
            "monthKey": _prev_m_end.strftime("%Y-%m"),
            "renewalReport": build_renewal_report(_prev_m_end.replace(day=1), _prev_m_end),
            "renewalReportRecalculadoEm": TODAY.isoformat(),
        }
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

        def sale_type_for(idm, sale_date_d, id_sale=None):
            # Pedido do usuario em 2026-09-25 (4a correcao do dia): alem de novo/renovacao/retorno,
            # tambem retorna se uma "renovacao" e ANTECIPADA — cliente comprou o novo contrato antes
            # do mes-calendario em que o contrato anterior realmente venceria (_endEfetivo em mes/ano
            # POSTERIOR ao mes/ano da nova venda). Regra validada nome-a-nome contra a planilha do
            # usuario (15 renovacoes de Piscina em setembro/2026: 8 normais + 7 antecipadas). So se
            # aplica a "renovacao"; None para "novo"/"retorno". Retorna (tipoVenda, renovacaoAntecipada).
            if not idm or not sale_date_d:
                return "novo", None, None
            # Ajuste manual pontual (pedido do usuario em 2026-09-29) — nao e regra geral.
            _override = _MANUAL_SALE_TYPE_OVERRIDES.get((unit["key"], idm, sale_date_d))
            if _override:
                return tuple(_override) + (False,) * (3 - len(_override))
            contracts = contracts_by_member.get(idm) or []
            # Pedido do usuario em 2026-09-29: plano CORTESIA e ignorado para classificar a venda —
            # compara sempre com o ultimo plano anterior a cortesia. Caso validado: Diana Cardoso
            # Lima Rodrigues (Fitness, idMember 3232) — Fitness Anual ate 27/02/2024, Cortesia 15 dias
            # 27/07-11/08/2026, nova venda 03/09/2026: antes contava como Renovacao (23 dias apos a
            # cortesia), agora conta como Retorno (2+ anos sem plano pago). Se o cliente so teve
            # cortesia antes, a venda conta como "novo".
            # Pedido do usuario em 2026-10-01: o proprio contrato desta venda nao pode entrar como
            # "anterior" — quando o inicio e retroativo (ex. Marina Rondon: vendido em 01/10 com inicio
            # 30/09) ele virava o "ultimo contrato" e a venda caia como Renovacao Antecipada.
            earlier = [c for c in contracts if c["_start"] and c["_start"] < sale_date_d and not is_transfer_record(c)
                       and not (id_sale is not None and c.get("idSale") == id_sale)
                       and "cortesia" not in _strip_accents((c.get("nameMembership") or "").strip().lower())]
            if not earlier:
                return "novo", None, None

            def end_for_gap(c):
                return c["_endEfetivo"] if c["_endEfetivo"] else datetime.date.max

            latest = max(earlier, key=end_for_gap)
            end_date = end_for_gap(latest)
            gap_days = (sale_date_d - end_date).days
            if gap_days > RETORNO_GAP_DAYS:
                return "retorno", None, None
            is_antecipada = end_date != datetime.date.max and (end_date.year, end_date.month) > (sale_date_d.year, sale_date_d.month)
            # Pedido do usuario em 2026-10-01: "Renovação Tardia" = renovou dentro dos 30 dias, mas num mes
            # POSTERIOR ao do vencimento (o contrato anterior venceu num mes e a venda caiu no seguinte).
            # A venda conta no mes em que foi feita; a renovacao conta no mes do vencimento.
            is_tardia = end_date != datetime.date.max and (end_date.year, end_date.month) < (sale_date_d.year, sale_date_d.month)
            return "renovacao", is_antecipada, is_tardia

        items = []
        total_value = 0.0
        for s in sales:
            if s.get("removed"):
                continue
            if s.get("idSale") in _MANUAL_IGNORE_SALE_IDS.get(unit["key"], set()):
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
            # Pedido do usuario em 2026-10-01: venda para quem ainda nao e cliente (ex. "PROFISSIONAL DE ALTA
            # PERMORMANCE", vendida a prospects) vinha sem nome — usa o nome do prospect.
            if not nome:
                _pr = s.get("prospect") or {}
                nome = " ".join(x for x in [_pr.get("firstName"), _pr.get("lastName")] if x).strip() or None
            tipo_venda, renovacao_antecipada, renovacao_tardia = sale_type_for(idm, sale_date_d, s.get("idSale"))
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
                    # Pedido do usuario em 2026-10-01: na unidade Fitness nada e classificado como Piscina —
                    # ex. item "MENSALIDADE" (Luanda Gabryeli, R$ 429, 06/09/2026) caia em "Piscina
                    # (recorrente)" pela palavra "mensal". No Fitness vira Fitness.
                    "plano": plano, "categoria": ("Fitness (Anual/Assinatura/Plus)" if (unit["key"] == "fitness" and categorize_sale_item(plano) == "Piscina (recorrente)") else categorize_sale_item(plano)), "valor": round(val, 2),
                    "dataVenda": sale_date, "consultor": e.get("consultor"), "tipoVenda": tipo_venda,
                    # Pedido do usuario em 2026-09-25 (4a correcao do dia): so preenchido (True/False)
                    # quando tipoVenda == "renovacao" — ver comentario em sale_type_for().
                    "renovacaoAntecipada": renovacao_antecipada,
                    "renovacaoTardia": renovacao_tardia,
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
                elif it["tipoVenda"] == "renovacao" and it.get("renovacaoTardia"):
                    tipo_label = "Renovação Tardia"
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

        # Pedido do usuario em 2026-09-30: "Vendas por semana" passa a usar a MESMA base dos tiles —
        # count = vendas de plano da unidade (Fitness+Bike no Fitness, Piscina recorrente na Piscina),
        # sem trocas de contrato (soma das semanas = "Numero de vendas"); valor = valor AJUSTADO (todas
        # as linhas, com adesao, sem trocas e sem festival — soma das semanas = "Valor ajustado").
        # Os numeros antigos (todas as linhas / valor bruto) ficam em countBruto/valorBruto.
        week_plan_cats = ("Fitness (Anual/Assinatura/Plus)", "Bike") if unit["key"] == "fitness" else ("Piscina (recorrente)",)
        by_week = []
        for seg_start, seg_end in week_segments:
            seg_count = 0
            seg_valor = 0.0
            seg_count_bruto = 0
            seg_valor_bruto = 0.0
            for it in items:
                dv = it["dataVenda"]
                try:
                    d = datetime.datetime.fromisoformat(dv).date() if dv else None
                except Exception:
                    d = None
                if d is None or not (seg_start <= d <= seg_end):
                    continue
                seg_count_bruto += 1
                seg_valor_bruto += it["valor"]
                if it["trocaContrato"]:
                    continue
                if it["categoria"] in week_plan_cats:
                    seg_count += 1
                if "festival" not in _strip_accents((it.get("plano") or "").lower()):
                    seg_valor += it["valor"]
            by_week.append({
                "weekStart": seg_start.isoformat(), "weekEnd": seg_end.isoformat(),
                "label": f"{seg_start.strftime('%d/%m')}–{seg_end.strftime('%d/%m')}",
                "count": seg_count, "valor": round(seg_valor, 2),
                "countBruto": seg_count_bruto, "valorBruto": round(seg_valor_bruto, 2),
            })

        # Pedido do usuario em 2026-09-29: as contagens Novo/Renovacao/Renovacao Antecipada/Retorno
        # (e o "Count" total usado no ticket medio) agora EXCLUEM vendas de troca de contrato — antes
        # so o VALOR era descontado (totalValueAjustado), a contagem continuava incluindo as trocas,
        # o que inflava "Renovações" e deflava o ticket medio (numerador ja descontado, denominador
        # nao). Os numeros "como se contasse tudo" (trocas incluidas) ficam disponiveis em campos
        # "*ComTrocas" separados, pra mostrar em letras miudas no dashboard — ver bucket_counts().
        def bucket_counts(lst):
            novos = [it for it in lst if it["tipoVenda"] == "novo"]
            renovacoes = [it for it in lst if it["tipoVenda"] == "renovacao" and not it.get("renovacaoAntecipada") and not it.get("renovacaoTardia")]
            renovacoes_antecipadas = [it for it in lst if it["tipoVenda"] == "renovacao" and it.get("renovacaoAntecipada")]
            renovacoes_tardias = [it for it in lst if it["tipoVenda"] == "renovacao" and it.get("renovacaoTardia")]
            retornos = [it for it in lst if it["tipoVenda"] == "retorno"]
            return {
                "count": len(lst), "valor": round(sum(it["valor"] for it in lst), 2),
                "novoCount": len(novos), "novoValue": round(sum(it["valor"] for it in novos), 2),
                "renovacaoCount": len(renovacoes), "renovacaoValue": round(sum(it["valor"] for it in renovacoes), 2),
                "renovacaoAntecipadaCount": len(renovacoes_antecipadas), "renovacaoAntecipadaValue": round(sum(it["valor"] for it in renovacoes_antecipadas), 2),
                "renovacaoTardiaCount": len(renovacoes_tardias), "renovacaoTardiaValue": round(sum(it["valor"] for it in renovacoes_tardias), 2),
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
                f"{prefix}RenovacaoTardiaCount": real["renovacaoTardiaCount"], f"{prefix}RenovacaoTardiaValue": real["renovacaoTardiaValue"],
                f"{prefix}RetornoCount": real["retornoCount"], f"{prefix}RetornoValue": real["retornoValue"],
                # "ComTrocas" = como ficaria se as vendas de troca de contrato NAO fossem excluidas
                # (fine print no dashboard) — so preenchido quando ha pelo menos 1 troca detectada
                # nessa categoria neste mes, senao e identico ao real (evita ruido visual a toa).
                f"{prefix}CountComTrocas": com_trocas["count"],
                f"{prefix}RenovacaoCountComTrocas": com_trocas["renovacaoCount"], f"{prefix}RenovacaoValueComTrocas": com_trocas["renovacaoValue"],
                f"{prefix}RenovacaoAntecipadaCountComTrocas": com_trocas["renovacaoAntecipadaCount"], f"{prefix}RenovacaoAntecipadaValueComTrocas": com_trocas["renovacaoAntecipadaValue"],
                f"{prefix}RenovacaoTardiaCountComTrocas": com_trocas["renovacaoTardiaCount"], f"{prefix}RenovacaoTardiaValueComTrocas": com_trocas["renovacaoTardiaValue"],
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
    troca_sale_ids_mes |= swapped_sale_ids

    # Pedido do usuario em 2026-10-01: venda feita no ultimo dia do mes DEPOIS da ultima coleta (ex.
    # Marcela Gama, 30/09) ficava de fora para sempre, e cancelamentos lancados depois tambem nao
    # entravam. Do dia 1 ao 5, o relatorio de vendas do MES ANTERIOR e recalculado inteiro e
    # regravado (sales_report_prev_month*.json). Usa as trocas de contrato daquele mes.
    if TODAY.day <= 5:
        _troca_cur = troca_sale_ids_mes
        try:
            _pv_end = month_start - datetime.timedelta(days=1)
            _pv_start = _pv_end.replace(day=1)
            _pv_auto = {it["idSaleNova"] for it in contract_swap_report["items"]
                        if _pv_start.isoformat() <= it["dataVendaNova"][:10] <= _pv_end.isoformat() and it.get("idSaleNova") is not None}
            troca_sale_ids_mes = _pv_auto | _MANUAL_TROCA_SALE_IDS.get(unit["key"], set()) | swapped_sale_ids
            _pv = build_sales_report(_pv_start, _pv_end)
            _pv_doc = dict(_pv)
            _pv_doc["monthKey"] = _pv_end.strftime("%Y-%m")
            _pv_auto_items = [m for m in contract_swap_report["items"] if m.get("idSaleNova") in _pv_auto]
            _pv_man = [it for it in _pv["items"] if it["idSale"] in troca_sale_ids_mes and it["idSale"] not in _pv_auto]
            _pv_doc["countTrocaContratoMes"] = len(_pv_auto_items) + len({it["idSale"] for it in _pv_man})
            _pv_doc["valorTrocaContratoMes"] = round(sum(m["valorVendaNova"] for m in _pv_auto_items) + sum(it["valor"] for it in _pv_man), 2)
            _pv_fest = [it for it in _pv["items"] if "festival" in _strip_accents((it.get("plano") or "").lower()) and it["idSale"] not in troca_sale_ids_mes]
            _pv_doc["countFestivalMes"] = len(_pv_fest)
            _pv_doc["valorFestivalMes"] = round(sum(it["valor"] for it in _pv_fest), 2)
            _pv_doc["totalValueAjustado"] = round(_pv["totalValue"] - _pv_doc["valorTrocaContratoMes"] - _pv_doc["valorFestivalMes"], 2)
            _pv_doc["recalculadoEm"] = TODAY.isoformat()
            _pv_doc["updatedAt"] = datetime.datetime.utcnow().isoformat() + "Z"
            SALES_PREV_OUT[unit["key"]] = _pv_doc
            print(f"[{unit['key']}] vendas do mes anterior recalculadas: total={_pv['total']} value={_pv['totalValue']}", file=sys.stderr)
        except Exception as e:
            print(f"[{unit['key']}] vendas do mes anterior falharam (resto segue normal): {e}", file=sys.stderr)
        finally:
            troca_sale_ids_mes = _troca_cur

    sales_report_this_month = build_sales_report(month_start, month_end)
    print(f"[{unit['key']}] sales this month: total={sales_report_this_month['total']} value={sales_report_this_month['totalValue']}", file=sys.stderr)

    loyalty_category = "Fitness (Anual/Assinatura/Plus)" if unit["key"] == "fitness" else "Piscina (recorrente)"
    loyalty_report = build_loyalty_report(loyalty_category, contracts_by_member, TODAY)
    loyalty_report["date"] = TODAY.isoformat()
    try:
        permanencia = build_permanencia(loyalty_category, contracts_by_member, TODAY)
        print(f"[{unit['key']}] permanencia: ativos={permanencia['resumo']['ativos']} ex={permanencia['resumo']['exClientes']}", file=sys.stderr)
    except Exception as e:
        print(f"[{unit['key']}] permanencia falhou (resto segue normal): {e}", file=sys.stderr)
        permanencia = None
    print(f"[{unit['key']}] fidelidade: fieis={loyalty_report['fieis']['count']} retornaram={loyalty_report['retornaram']['count']} parcelas={loyalty_report['parcelas']['totalClients']} (media R$ {loyalty_report['parcelas']['avgParcela']:.2f})", file=sys.stderr)

    sales_report_this_month_doc = dict(sales_report_this_month)
    sales_report_this_month_doc["monthKey"] = TODAY.strftime("%Y-%m")
    _manual_troca_items = [it for it in sales_report_this_month["items"]
                           if it["idSale"] in troca_sale_ids_mes and it["idSale"] not in troca_auto_ids_mes]
    _manual_troca_count = len({it["idSale"] for it in _manual_troca_items})
    _manual_troca_value = round(sum(it["valor"] for it in _manual_troca_items), 2)
    sales_report_this_month_doc["countTrocaContratoMes"] = contract_swap_report["countMesCorrente"] + _manual_troca_count
    sales_report_this_month_doc["valorTrocaContratoMes"] = round(contract_swap_report["valorMesCorrente"] + _manual_troca_value, 2)
    # Pedido do usuario em 2026-09-30: vendas de FESTIVAL (ex. "FESTIVAL TRAVESSIA", R$ 60 cada, sem
    # cliente) nao sao venda de plano — continuam no "Valor vendido" e em "Vendas por tipo de
    # contrato", mas saem do "Valor ajustado" (e portanto do ticket medio). Adesao continua no
    # valor ajustado, como antes. So a palavra "festival" — nao expandir sem o usuario pedir.
    _festival_items = [it for it in sales_report_this_month["items"]
                       if "festival" in _strip_accents((it.get("plano") or "").lower()) and it["idSale"] not in troca_sale_ids_mes]
    sales_report_this_month_doc["countFestivalMes"] = len(_festival_items)
    sales_report_this_month_doc["valorFestivalMes"] = round(sum(it["valor"] for it in _festival_items), 2)
    sales_report_this_month_doc["totalValueAjustado"] = round(sales_report_this_month["totalValue"] - sales_report_this_month_doc["valorTrocaContratoMes"] - sales_report_this_month_doc["valorFestivalMes"], 2)
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
    # Pedido do usuario em 2026-09-30 (2a parte): "Valor perdido" = SALDO que faltava pagar de cada
    # contrato cancelado no mes (remainingValue do contrato no EVO), nao o valor cheio do contrato.
    # Ficam fora (saldo 0): trocas internas de plano (motivo) e cancelamentos de clientes que tiveram
    # venda de TROCA DE CONTRATO no mes (ex. Leticia Bonfim, Yasmin Silva — motivo no EVO nao diz
    # "troca"). Contrato antigo baixado sem saldo (ex. Antonio Galrao, contrato 2022-2023) ja da 0.
    # Se o contrato nao for encontrado, cai no valor do contrato. O valor cheio (sem trocas) fica em
    # cancelledValueContrato e o bruto (com trocas) em cancelledValueBruto.
    _troca_members = {it["idMember"] for it in sales_report_this_month["items"] if it["idSale"] in troca_sale_ids_mes}
    cancelled_value_contrato = cancelled_value
    cancelled_value = 0.0
    # Pedido do usuario em 2026-09-30 (3a parte): esses casos tambem saem da CONTAGEM de Cancelamentos
    # (e do Churn) — cancelamentos_excluidos e descontado no dashboard, junto com as trocas internas:
    # (a) cliente com venda de troca de contrato no mes; (b) baixa administrativa de contrato que ja
    # tinha terminado antes do inicio do mes (ex. Antonio Galrao, contrato 2022-2023 baixado em 09/2026).
    cancelamentos_excluidos = 0
    for r in cancelled:
        try:
            _idm = int(r.get("IdCliente"))
        except Exception:
            _idm = None
        if categorize_reason(r.get("MotivoCancelamento")) == "Troca interna de plano":
            r["_saldo"], r["_obs"] = 0.0, "troca interna de plano"
            continue
        if _idm in _troca_members:
            r["_saldo"], r["_obs"] = 0.0, "cliente com troca de contrato no mês"
            cancelamentos_excluidos += 1
            continue
        _dt = parse_date_br(str(r.get("DtCancelamento") or "")[:10])
        _val = parse_value_br(r.get("Valor"))
        _cands = [c for c in contracts_by_member.get(_idm, []) if _dt and c.get("_cancelDate") == _dt]
        if _cands:
            _c = min(_cands, key=lambda c: abs((c.get("saleValue") or 0) - _val))
            if _c.get("_end") and _c["_end"] < month_start:
                r["_saldo"], r["_obs"] = 0.0, "baixa administrativa — contrato já tinha terminado"
                cancelamentos_excluidos += 1
                continue
            r["_saldo"], r["_obs"] = round(_c.get("remainingValue") or 0.0, 2), None
        else:
            r["_saldo"], r["_obs"] = round(_val, 2), "saldo não encontrado — valor do contrato"
        cancelled_value += r["_saldo"]
    print(f"[{unit['key']}] valor perdido (saldo): R$ {cancelled_value:.2f} (valor cheio sem trocas R$ {cancelled_value_contrato:.2f})", file=sys.stderr)

    try:
        client_lists = build_client_lists(TODAY, active, all_members_full, contracts_by_member, current_contract_for,
                                          vip_ids, suspended_ids, all_time_prospects, cancelled, prospects)
        if permanencia:
            client_lists["permanencia"] = permanencia
    except Exception as e:
        print(f"[{unit['key']}] listas de clientes falharam (resto segue normal): {e}", file=sys.stderr)
        client_lists = None

    client_report_this_month_doc = {
        "monthKey": TODAY.strftime("%Y-%m"),
        "cancelledCount": len(cancelled), "internalPlanChanges": internal_plan_changes,
        "cancelamentosExcluidos": cancelamentos_excluidos, "cancelledValue": round(cancelled_value, 2), "cancelledValueContrato": round(cancelled_value_contrato, 2), "cancelledValueBruto": round(cancelled_value_bruto, 2), "cancelReasons": reasons,
        "newLeads": total_leads, "convertedLeads": len(converted),
        "conversionRate": round(len(converted) / total_leads * 100, 1) if total_leads else 0,
        "renewalReport": renewal_report_this_month,
        "activeClients": len(active), "vipClients": vip_count, "vipByContractType": vip_by_contract_type,
        "suspendedClients": suspended_count, "exClients": ex_clients_count, "opportunities": opportunities_count,
        "totalClients": len(active) + suspended_count,
        "updatedAt": datetime.datetime.utcnow().isoformat() + "Z",
    }
    # Pedido do usuario em 2026-09-30: base do Churn = ativos + suspensos no dia 1o do mes, guardada
    # PARA SEMPRE no documento mensal. So e gravada na coleta do dia 1o; nos outros dias o campo nao
    # vai no documento e a rotina grava clientReports/clientReportsPiscina com "update" (merge), entao
    # o valor do dia 1o fica preservado. Setembro/2026 foi gravado a mao (Fitness 602, Piscina 335).
    if TODAY.day == 1:
        client_report_this_month_doc["baseInicioMes"] = len(active) + suspended_count
    if client_lists is not None:
        client_report_this_month_doc["cancelledList"] = client_lists["cancelados"]
        client_report_this_month_doc["leadsList"] = client_lists["leads"]

    client_snapshot = {
        "date": TODAY.isoformat(), "monthKey": TODAY.strftime("%Y-%m"), "nextMonthKey": next_month_start.strftime("%Y-%m"),
        "activeClients": len(active), "cancelledMTD": len(cancelled), "internalPlanChangesMTD": internal_plan_changes,
        "cancelledValueMTD": round(cancelled_value, 2), "cancelamentosExcluidosMTD": cancelamentos_excluidos, "newLeadsMTD": total_leads, "convertedLeadsMTD": len(converted),
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

    try:
        _names = {m["idMember"]: (" ".join(x for x in [m.get("firstName"), m.get("lastName")] if x).strip() or None) for m in all_members_full}
        _vip_set, _susp_set = set(vip_ids), set(suspended_ids)

        def _status_for(idm):
            if not idm:
                return "Inativo"
            if idm in active_ids or idm in _vip_set:
                return "Ativo"
            if idm in _susp_set:
                return "Suspenso"
            return "Inativo"
        RECEIVABLES_OUT[unit["key"]] = build_receivables_report(api_get_raw, month_start, month_end, TODAY, active_ids, _names, current_contract_for, _status_for, access_stats)
        RECEIVABLES_OUT[unit["key"]]["inadimplentesAcumulado"] = build_delinquency_report(api_get_raw, TODAY, active_ids, _names, current_contract_for, _status_for, access_stats)
        try:
            RECEIVABLES_OUT[unit["key"]]["caixa"] = build_cash_report(api_get_raw, month_start, month_end, TODAY)
        except Exception as e:
            print(f"[{unit['key']}] caixa do mes falhou (resto segue normal): {e}", file=sys.stderr)
        # Pedido do usuario em 2026-09-30 (6a parte): projecao do MES SEGUINTE — tudo que ja esta
        # lancado no EVO com vencimento no proximo mes (parcelas de cartao ja autorizadas = garantido;
        # boleto/saldo/pix = a receber). Refeito todo dia ate o mes virar o corrente.
        RECEIVABLES_OUT[unit["key"] + ":next"] = build_receivables_report(api_get_raw, next_month_start, next_month_end, TODAY, active_ids, _names, current_contract_for, _status_for, access_stats)
        RECEIVABLES_OUT[unit["key"] + ":next"]["projecao"] = True
        if TODAY.day == 1:
            _prev_end = month_start - datetime.timedelta(days=1)
            RECEIVABLES_OUT[unit["key"] + ":prev"] = build_receivables_report(api_get_raw, _prev_end.replace(day=1), _prev_end, TODAY, active_ids, _names, current_contract_for, _status_for, access_stats)
            # fechamento do mes anterior guarda o mesmo acumulado de hoje (= tudo vencido ate o ultimo dia dele)
            RECEIVABLES_OUT[unit["key"] + ":prev"]["inadimplentesAcumulado"] = RECEIVABLES_OUT[unit["key"]]["inadimplentesAcumulado"]
        # Pedido do usuario em 2026-10-01: caixa do MES ANTERIOR (fechamento). No dia 1o vai junto do
        # fechamento acima; dos dias 2 a 5 e regravado so o campo "caixa" (update), pegando creditos
        # e baixas lancados com atraso.
        _pe = month_start - datetime.timedelta(days=1)
        if TODAY.day <= 5:
            try:
                _cx = build_cash_report(api_get_raw, _pe.replace(day=1), _pe, TODAY)
                if TODAY.day == 1 and (unit["key"] + ":prev") in RECEIVABLES_OUT:
                    RECEIVABLES_OUT[unit["key"] + ":prev"]["caixa"] = _cx
                elif TODAY.day != 1:
                    RECEIVABLES_OUT[unit["key"] + ":caixaPrev"] = {"monthKey": _pe.strftime("%Y-%m"), "caixa": _cx}
            except Exception as e:
                print(f"[{unit['key']}] caixa do mes anterior falhou (resto segue normal): {e}", file=sys.stderr)
        _rr = RECEIVABLES_OUT[unit["key"]]
        print(f"[{unit['key']}] inadimplentes acumulados: {_rr['inadimplentesAcumulado']['qtdPessoas']} pessoas, {_rr['inadimplentesAcumulado']['qtdParcelas']} parcelas, R$ {_rr['inadimplentesAcumulado']['total']:.2f}", file=sys.stderr)
        print(f"[{unit['key']}] recebimentos: {_rr['qtdPessoas']} pessoas, total R$ {_rr['totalValor']:.2f}, pago R$ {_rr['totalPago']:.2f}, inadimplente R$ {_rr['totalInadimplente']:.2f} ({_rr['qtdPessoasInadimplentes']} pessoas), ativos sem recebivel {_rr['qtdAtivosSemRecebivel']}", file=sys.stderr)
    except Exception as e:
        print(f"[{unit['key']}] recebimentos falharam (resto segue normal): {e}", file=sys.stderr)

    # v7 (pedido do usuario em 2026-11-11): parcelas, baixas e consultor para o resumo do mes. So acrescenta:
    # nao altera member_cache nem os contratos, e uma falha aqui nao afeta nada do v6.
    if not os.environ.get("COLETA_DATA_REF"):
        try:
            EXTRAS_OUT[unit["key"]] = build_month_extras(unit, api_get_raw, API_BASE_V2, TODAY, month_start, month_end,
                                                         next_month_start, next_month_end, all_members_full, all_contracts,
                                                         renovacao=renewal_report_this_month)
        except Exception as e:
            EXTRAS_ERR[unit["key"]] = repr(e)
            print(f"[{unit['key']}] v7 extras falharam (resto segue normal): {e!r}", file=sys.stderr)
        # v7.1 (pedido do usuario em 2026-11-11): dias 1 a 5, refaz tambem o MES ANTERIOR com a posicao final
        # (pagamentos ate o ultimo dia), para o fechamento comercial do 3o dia util.
        if TODAY.day <= 5:
            _pe = month_start - datetime.timedelta(days=1)
            try:
                EXTRAS_PREV_OUT[unit["key"]] = build_month_extras(unit, api_get_raw, API_BASE_V2, TODAY, _pe.replace(day=1), _pe,
                                                                  month_start, month_end, all_members_full, all_contracts,
                                                                  class_ref=month_start, buscar_telefone=False,
                                                                  renovacao=(CLIENT_PREV_OUT.get(unit["key"]) or {}).get("renewalReport"))
            except Exception as e:
                EXTRAS_PREV_ERR[unit["key"]] = repr(e)
                print(f"[{unit['key']}] v7.1 fechamento do mes anterior falhou (resto segue normal): {e!r}", file=sys.stderr)

    workout_snapshot = None
    frequencia_report = None

    if unit["key"] == "piscina":
        freq_list = []
        turma_matched = 0
        RECENT_TURMA_DAYS = 14
        recent_turma_cutoff = (TODAY - datetime.timedelta(days=RECENT_TURMA_DAYS - 1)).isoformat()
        # Pedido do usuario em 2026-10-05: 7 dias completos anteriores a coleta (D-7 a D-1).
        window7_cutoff = (TODAY - datetime.timedelta(days=7)).isoformat()
        window7_end = TODAY.isoformat()
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
                within7 = [r for r in records if window7_cutoff <= (r.get("date") or "")[:10] < window7_end]
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
            "janelaInicio": window7_cutoff,
            "janelaFim": (TODAY - datetime.timedelta(days=1)).isoformat(),
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
            # Pedido do usuario em 2026-10-06: frequencia de TODOS os alunos (nao so os de "Contratos
            # vencendo") para a "Frequencia por carteira de professor" do dashboard — dias com entrada
            # na catraca/facial nos ultimos 7 e 30 dias (mesma base do acessos30d da renovacao).
            "acessos7d": sum(1 for d in (entry_dates_by_member.get(idc) or set()) if d >= _cut7),
            "acessos30d": access_stats(idc)[1],
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
                "acessos7d": c.get("acessos7d"), "acessos30d": c.get("acessos30d"),
            }
            for c in sorted(out, key=lambda x: (x["nome"] or ""))
        ],
        "updatedAt": datetime.datetime.utcnow().isoformat() + "Z", "errors": errors,
    }

    return client_snapshot, workout_snapshot, sales_report_this_month_doc, client_report_this_month_doc, loyalty_report, frequencia_report, client_lists


if __name__ == "__main__":
    # v7: o laco do v6 fica igual, so dentro de try/finally — no fim (com sucesso ou erro) grava os extras
    try:
        for unit in UNITS:
            snap, workout, sales_doc, client_doc, loyalty_doc, frequencia_doc, lists_doc = build_for_unit(unit)
            suf = unit["suffix"]
            if os.environ.get("COLETA_DATA_REF"):
                _keys = ["monthKey", "cancelledCount", "internalPlanChanges", "cancelamentosExcluidos", "cancelledValue",
                         "cancelledValueContrato", "cancelledValueBruto", "cancelReasons", "newLeads", "convertedLeads",
                         "conversionRate", "renewalReport", "cancelledList", "leadsList"]
                _bf = {k: client_doc[k] for k in _keys if k in client_doc}
                _bf["recalculadoEm"] = datetime.date.today().isoformat()
                with open(f"client_report_backfill{suf and '_' + suf.lower()}.json", "w") as f:
                    json.dump(_bf, f, ensure_ascii=False, indent=2)
                # Pedido do usuario em 2026-11-09: o recalculo de mes fechado tambem grava vendas e
                # recebimentos daquele mes (para preencher o historico de jan-set/2026 no painel novo).
                # Recebimentos refletem o status ATUAL de cada parcela (pagas depois do mes aparecem pagas).
                _sv = dict(sales_doc)
                _sv["recalculadoEm"] = datetime.date.today().isoformat()
                with open(f"sales_report_backfill{suf and '_' + suf.lower()}.json", "w") as f:
                    json.dump(_sv, f, ensure_ascii=False, indent=2)
                if unit["key"] in RECEIVABLES_OUT:
                    _rv = dict(RECEIVABLES_OUT[unit["key"]])
                    _rv["recalculadoEm"] = datetime.date.today().isoformat()
                    with open(f"receivables_report_backfill{suf and '_' + suf.lower()}.json", "w") as f:
                        json.dump(_rv, f, ensure_ascii=False)
                else:
                    print(f"[{unit['key']}] recebimentos do backfill indisponiveis (ver erro acima)", file=sys.stderr)
                print(f"[{unit['key']}] done. backfill {_bf['monthKey']} (clientes, vendas, recebimentos)", file=sys.stderr)
                continue
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
            if unit["key"] in RECEIVABLES_OUT:
                with open(f"receivables_report_month{suf and '_' + suf.lower()}.json", "w") as f:
                    json.dump(RECEIVABLES_OUT[unit["key"]], f, ensure_ascii=False)
            if unit["key"] + ":next" in RECEIVABLES_OUT:
                with open(f"receivables_report_next_month{suf and '_' + suf.lower()}.json", "w") as f:
                    json.dump(RECEIVABLES_OUT[unit["key"] + ":next"], f, ensure_ascii=False)
            if unit["key"] in SALES_PREV_OUT:
                with open(f"sales_report_prev_month{suf and '_' + suf.lower()}.json", "w") as f:
                    json.dump(SALES_PREV_OUT[unit["key"]], f, ensure_ascii=False, indent=2)
            if unit["key"] in CLIENT_PREV_OUT:
                with open(f"client_report_prev_month{suf and '_' + suf.lower()}.json", "w") as f:
                    json.dump(CLIENT_PREV_OUT[unit["key"]], f, ensure_ascii=False, indent=2)
            if unit["key"] + ":caixaPrev" in RECEIVABLES_OUT:
                with open(f"receivables_caixa_prev_month{suf and '_' + suf.lower()}.json", "w") as f:
                    json.dump(RECEIVABLES_OUT[unit["key"] + ":caixaPrev"], f, ensure_ascii=False)
            if unit["key"] + ":prev" in RECEIVABLES_OUT:
                with open(f"receivables_report_prev_month{suf and '_' + suf.lower()}.json", "w") as f:
                    json.dump(RECEIVABLES_OUT[unit["key"] + ":prev"], f, ensure_ascii=False)
            extra_size = len(json.dumps(workout)) if workout is not None else len(json.dumps(frequencia_doc))
            extra_label = "workout" if workout is not None else "frequencia"
            print(f"[{unit['key']}] done. snapshot size={len(json.dumps(snap))} {extra_label} size={extra_size} loyalty size={len(json.dumps(loyalty_doc))}", file=sys.stderr)
    finally:
        finalize_extras()
