"""Metadata-only per-user analyses. Log hashes are NOT licence user keys."""
from __future__ import annotations

import base64
import csv
import io
import re
from collections import Counter, defaultdict
from datetime import timedelta

from .log_taxonomy import classify, load_taxonomy

PERIODS = {"7d": 7, "30d": 30, "90d": 90, "365d": 365, "all": None}
HOSTS = {"ALL", "WD", "EX", "OL"}


def _filters(store, period, host):
    if period not in PERIODS or host not in HOSTS:
        raise ValueError('Invalid user-analytics period or Office host.')
    days = PERIODS[period]
    params = []
    clauses = ['invoked=1']
    if days:
        clauses.extend(('date>=?', 'date<=?'))
        params.extend(((store.clock().date() - timedelta(days=days-1)).isoformat(), store.clock().date().isoformat()))
    if host != 'ALL':
        clauses.append('host_id=?')
        params.append(host)
    return ' WHERE ' + ' AND '.join(clauses), params


def _bucket(calls: int):
    if calls == 1: return '1'
    if calls <= 5: return '>1–5'
    if calls <= 20: return '>5–20'
    if calls <= 50: return '>20–50'
    return '>50'


def function_users(store, period='30d', host='ALL', selected=None, *, all_rows=False):
    from datetime import date
    from statistics import median
    from .usage_semantics import describe, primary_family

    where, params = _filters(store, period, host)
    if selected is not None and not re.fullmatch(r'[0-9a-fA-F]{8,128}', selected):
        raise ValueError('Invalid pseudonymous log user identifier.')
    selected = selected.upper() if selected else None
    taxonomy = load_taxonomy(store.settings.data_dir)
    def user_stats():
        return {'hosts':Counter(),'features':Counter(),'families':Counter(),'days':Counter(),
                'tools':Counter(),'agentic':0,'standard':0}
    data = defaultdict(user_stats)
    with store.connect() as con:
        for r in con.execute('SELECT user_hash,host_id,feature,date,COUNT(*) n FROM log_events'+where+
                ' GROUP BY user_hash,host_id,feature,date ORDER BY user_hash,date,feature,host_id',params):
            e = describe(r['feature'])
            n = r['n']
            if e.kind in ('tool','skill','agent'):
                data[r['user_hash']]['tools'][e.label] += n
            if not e.primary:
                continue
            u = data[r['user_hash']]
            u['hosts'][r['host_id']] += n
            u['features'][e.label] += n
            u['families'][primary_family(e,taxonomy)['label']] += n
            u['days'][r['date']] += n
            u['agentic' if e.kind == 'run' else 'standard'] += n
    keys = sorted((k for k,v in data.items() if v['days']), key=lambda k:(-sum(data[k]['days'].values()),k))
    days = PERIODS[period]
    if days is None and keys:
        first = min(min(data[k]['days']) for k in keys)
        last = max(max(data[k]['days']) for k in keys)
        days = (date.fromisoformat(last)-date.fromisoformat(first)).days+1
    total = 0
    listing = []
    distribution = Counter()
    selected_detail = None
    for pos,key in enumerate(keys):
        u = data[key]
        events = u['days']; active = len(events); n=sum(events.values()); total+=n
        weeks = Counter()
        for day,value in events.items():
            d = date.fromisoformat(day); monday = (d-timedelta(days=d.weekday())).isoformat()
            weeks[monday] += value
        distribution[_bucket(n/active)] += 1
        if not all_rows and pos>=200 and key!=selected:
            continue
        metrics = dict(calls_per_active_day=round(n/active,2), calls_per_calendar_day=round(n/days,2),
                    calls_per_7_days=round(7*n/days,2),calls_per_active_week=round(n/len(weeks),2),
                    active_weeks=len(weeks), active_days_per_active_week=round(active/len(weeks),2),
                    median_per_active_day=median(events.values()),peak_day_count=max(events.values()))
        row = dict(user_hash=key,invocations=n,active_days=active,first_day=min(events),last_day=max(events),
                   agentic=u['agentic'],standard=u['standard'],hosts=dict(u['hosts']),
                   distinct_features=len(u['features']),agentic_share_pct=round(100*u['agentic']/n,1),**metrics)
        if all_rows or pos<200: listing.append(row)
        if key==selected:
            selected_detail = dict(row,features=[dict(feature=k,count=v) for k,v in u['features'].most_common(100)],
                                   families=[dict(family=k,count=v) for k,v in u['families'].most_common()],
                                   tool_details=[dict(feature=k,count=v) for k,v in u['tools'].most_common(30)],
                                   days=[dict(date=k,count=v) for k,v in sorted(events.items())],
                                   weeks=[dict(week_start=k,count=v) for k,v in sorted(weeks.items())])
    return dict(period=period,host=host,identity_basis='SharedLogger username hash, not a LicenseCounter key.',
            total_users=len(keys),total_invocations=total,users=listing,selected=selected_detail,period_days=days,
            distribution=[dict(range=k,users=distribution[k]) for k in ('1','>1–5','>5–20','>20–50','>50')],
            distribution_basis='Mean primary events per active log day, bucketed before rounding.',
            coverage='Observed direct commands and explicit run starts. Missing evidence is not proof of inactivity.')


def reveal_obfuscated_identity(value):
    """Display the documented o1_ reversible client encoding only when anon=false."""
    if not value or not value.startswith('o1_'):
        return None
    token = value[3:]
    if not re.fullmatch(r'[a-zA-Z0-9_-]{1,512}', token):
        return None
    try:
        name = base64.urlsafe_b64decode(token + '=' * (-len(token) % 4)).decode('utf-8')
        return name if name.isprintable() and len(name) <= 256 else None
    except (ValueError, UnicodeError):
        return None


def license_users(store, month):
    from .reports import valid_month
    valid_month(month)
    with store.connect() as conn:
        rows = conn.execute('''SELECT b.user_key,b.network_key,b.network_scope,b.product_id,
                MIN(b.anonymous) AS anonymous,MAX(b.user_obf) AS user_obf,
                GROUP_CONCAT(DISTINCT h.host_id) AS hosts,
                COUNT(DISTINCT d.loaded_date) AS loaded_days,
                MIN(b.first_received_utc) AS first_received_utc,MAX(b.last_received_utc) AS last_received_utc
                FROM billing_usage b LEFT JOIN billing_hosts h USING(billing_event_id)
                LEFT JOIN daily_usage d USING(billing_event_id) WHERE b.month=?
                GROUP BY b.user_key,b.network_key,b.product_id ORDER BY b.network_scope,b.user_key LIMIT 1000''',(month,)).fetchall()
    result = []
    for row in rows:
        obj = dict(row)
        obj['display_name'] = reveal_obfuscated_identity(obj['user_obf']) if not obj['anonymous'] else None
        obj.pop('user_obf',None)
        obj['anonymous'] = bool(obj['anonymous'])
        result.append(obj)
    return {'month':month,'users':result,'identity_basis':'LicenseCounter user key. Not correlated with SharedLogger username hash.',
            'warning':'Anon=false may contain a reversible configured username. Do not assume these users correspond to LogPath hashes.'}


def _csv_cell(value):
    """Protect Excel from CSV formula injection for arbitrary identity/source text."""
    value = '' if value is None else str(value)
    if value.lstrip().startswith(('=', '+', '-', '@', '\t', '\r', '\n')):
        return "'" + value
    return value


def _csv(rows):
    stream = io.StringIO(newline='')
    writer = csv.writer(stream, lineterminator='\r\n')
    for row in rows:
        writer.writerow([_csv_cell(cell) for cell in row])
    return '\ufeff' + stream.getvalue()


def user_summary_csv(data):
    rows = [['Log user hash', 'Invocations', 'Active log days', 'Average invocations per observed day',
             'Agent run starts', 'Direct commands', 'Word', 'Excel', 'Outlook', 'Distinct functions',
             'Observed per calendar day','Observed per 7 days','Per active ISO week','Active ISO weeks','Active days per active week','Median per active day','Peak day count','Denominator days']]
    for u in data['users']:
        rows.append([u['user_hash'],u['invocations'],u['active_days'],u['calls_per_active_day'],
                     u['agentic'],u['standard'],u['hosts'].get('WD',0),u['hosts'].get('EX',0),
                     u['hosts'].get('OL',0),u['distinct_features'],u['calls_per_calendar_day'],u['calls_per_7_days'],
                     u['calls_per_active_week'],u['active_weeks'],u['active_days_per_active_week'],u['median_per_active_day'],u['peak_day_count'],data['period_days']])
    return _csv(rows)


def license_users_csv(data):
    rows = [['Month','User key','Username (if Anon=False and available)','License scope','Scope key',
             'Product','Office hosts','Observed load days','Anonymous']]
    for u in data['users']:
        rows.append([data['month'],u['user_key'],u['display_name'],u['network_scope'],u['network_key'],
                     u['product_id'],u['hosts'],u['loaded_days'],u['anonymous']])
    return _csv(rows)
