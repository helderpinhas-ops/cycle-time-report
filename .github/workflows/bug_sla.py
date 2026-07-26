"""
Bug SLA Report — CAECS, GTM, DTAL, DTK
=======================================
Measures resolution time for Urgent and High priority bugs
against SLA targets:
  - Urgent: 90% resolved within 6 business hours
  - High:   75% resolved within 1.5 business days (12 business hours)

Only counts bugs with label 'jira_escalated' OR 'kanban'.
Uses business hours only (Mon–Fri, excluding PT public holidays).

Requirements:
    pip install requests python-dateutil openpyxl

Environment variables:
    JIRA_EMAIL         your Atlassian login email
    JIRA_API_TOKEN     Atlassian API token
    SLACK_WEBHOOK_URL  Slack webhook for #dtge_v2mom_sync (optional)
"""

import os
import requests
from datetime import datetime, timezone, timedelta
from dateutil import parser as dp
from collections import defaultdict

# ── Configuration ──────────────────────────────────────────────────────────────
JIRA_BASE   = "https://outsystemsrd.atlassian.net/rest/api/3"
JIRA_BROWSE = "https://outsystemsrd.atlassian.net/browse"
JIRA_EMAIL  = os.environ["JIRA_EMAIL"]
JIRA_TOKEN  = os.environ["JIRA_API_TOKEN"]
SLACK_URL   = os.environ.get("SLACK_WEBHOOK_URL")

AUTH    = (JIRA_EMAIL, JIRA_TOKEN)
HEADERS = {"Accept": "application/json"}

# ── SLA Targets ────────────────────────────────────────────────────────────────
SLA_URGENT_BH  = 6.0    # business hours — 90% target
SLA_URGENT_PCT = 90
SLA_HIGH_BH    = 12.0   # business hours (1.5 days) — 75% target
SLA_HIGH_PCT   = 75

# ── JQLs per project (bugs only, jira_escalated OR kanban label) ──────────────
# Looking back 12 weeks to match cycle time report window
LOOKBACK_DAYS = 84  # 12 weeks

PROJECT_JQLS = {
    "CAECS": (
        'project = CAECS '
        'AND issuetype = Bug '
        'AND (labels = jira_escalated OR labels = kanban) '
        'AND status != Discarded '
        'AND (Sprint not in openSprints() OR Sprint is EMPTY) '
        'AND created >= -{days}d '
        'ORDER BY created DESC'
    ),
    "DTAL": (
        'project = DTAL '
        'AND issuetype = Bug '
        'AND (labels = jira_escalated OR labels = kanban) '
        'AND status != Discarded '
        'AND (Sprint not in openSprints() OR Sprint is EMPTY) '
        'AND issuetype not in (Sub-task, "Bug - Sub-Task", "Technical Sub-Task") '
        'AND created >= -{days}d '
        'ORDER BY created DESC'
    ),
    "GTM": (
        'project = GTM '
        'AND issuetype = Bug '
        'AND (labels = jira_escalated OR labels = kanban) '
        'AND status != Discarded '
        'AND created >= -{days}d '
        'ORDER BY created DESC'
    ),
    "DTK": (
        'project = "Developers Enablement" '
        'AND issuetype = Bug '
        'AND (labels = jira_escalated OR labels = kanban) '
        'AND status != Discarded '
        'AND (Sprint not in openSprints() OR Sprint is EMPTY) '
        'AND created >= -{days}d '
        'ORDER BY created DESC'
    ),
}

# ── Portuguese Public Holidays ─────────────────────────────────────────────────
_holiday_cache = {}

def easter_date(year):
    a = year % 19; b = year // 100; c = year % 100
    d = b // 4;    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4;    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    month = (h + l - 7 * m + 114) // 31
    day   = (h + l - 7 * m + 114) % 31 + 1
    return datetime(year, month, day, tzinfo=timezone.utc)

def pt_holidays(year):
    if year in _holiday_cache:
        return _holiday_cache[year]
    easter      = easter_date(year)
    good_friday = easter - timedelta(days=2)
    corpus      = easter + timedelta(days=60)
    fmt = lambda d: d.strftime("%Y-%m-%d")
    holidays = {
        f"{year}-01-01", f"{year}-04-25", f"{year}-05-01",
        f"{year}-06-10", f"{year}-08-15", f"{year}-10-05",
        f"{year}-11-01", f"{year}-12-01", f"{year}-12-08",
        f"{year}-12-25", fmt(good_friday), fmt(corpus),
    }
    _holiday_cache[year] = holidays
    return holidays

def is_business_day(dt):
    if dt.weekday() >= 5:
        return False
    return dt.strftime("%Y-%m-%d") not in pt_holidays(dt.year)

def business_hours_between(start_dt, end_dt):
    """
    Count business hours between two UTC datetimes.
    Business hours = Mon-Fri 09:00-18:00 Lisbon time (approx UTC+0/+1).
    We use a simplified model: full working days × 8h + partial hours.
    """
    if end_dt <= start_dt:
        return 0.0

    BH_START = 9   # 09:00
    BH_END   = 17  # 17:00 (8h day)
    BH_HOURS = BH_END - BH_START

    total_bh = 0.0
    cursor   = datetime(start_dt.year, start_dt.month, start_dt.day,
                        tzinfo=timezone.utc)

    while cursor < end_dt:
        if is_business_day(cursor):
            day_bh_start = cursor.replace(hour=BH_START, minute=0, second=0, microsecond=0)
            day_bh_end   = cursor.replace(hour=BH_END,   minute=0, second=0, microsecond=0)

            # Clamp to [start_dt, end_dt]
            effective_start = max(day_bh_start, start_dt)
            effective_end   = min(day_bh_end,   end_dt)

            if effective_end > effective_start:
                total_bh += (effective_end - effective_start).total_seconds() / 3600

        cursor += timedelta(days=1)

    return total_bh

# ── Jira helpers ───────────────────────────────────────────────────────────────
def jira_get(path, params=None):
    r = requests.get(
        f"{JIRA_BASE}/{path}",
        params=params, auth=AUTH, headers=HEADERS, timeout=30
    )
    r.raise_for_status()
    return r.json()

def get_bugs(project, jql):
    """Fetch all bugs for a project using paginated search/jql."""
    issues, page_token, page = [], None, 1
    while True:
        params = {
            "jql":        jql,
            "maxResults": 100,
            "fields":     "summary,priority,status,created,resolutiondate,labels",
        }
        if page_token:
            params["nextPageToken"] = page_token
        data  = jira_get("search/jql", params)
        batch = data.get("issues", [])
        issues.extend(batch)
        if data.get("isLast", True) or not batch:
            break
        page_token = data.get("nextPageToken")
        if not page_token:
            break
        page += 1
    return issues

# ── SLA Calculation ────────────────────────────────────────────────────────────
def compute_bug_sla(issues):
    """
    For each resolved bug, compute business hours from created → resolved.
    For open bugs, compute business hours from created → now (in-flight).
    Returns per-priority stats.
    """
    now = datetime.now(timezone.utc)
    results = []

    for issue in issues:
        fields   = issue["fields"]
        priority = (fields.get("priority") or {}).get("name", "Normal").lower()
        status   = fields.get("status", {}).get("name", "").lower()
        created  = dp.parse(fields["created"]).astimezone(timezone.utc)
        resolved = dp.parse(fields["resolutiondate"]).astimezone(timezone.utc) \
                   if fields.get("resolutiondate") else None
        is_done  = status in ("done", "closed", "resolved", "ready for production")

        end_time  = resolved if (is_done and resolved) else now
        bh        = business_hours_between(created, end_time)
        in_flight = not is_done

        results.append({
            "key":       issue["key"],
            "summary":   fields.get("summary", ""),
            "priority":  priority,
            "status":    fields.get("status", {}).get("name", ""),
            "bh":        round(bh, 1),
            "is_done":   is_done,
            "in_flight": in_flight,
            "created":   created.strftime("%Y-%m-%d"),
            "resolved":  resolved.strftime("%Y-%m-%d") if resolved else "open",
        })

    return results

def sla_stats(bugs, priority_filter, sla_hours, sla_pct_target):
    """Compute SLA compliance for a given priority."""
    filtered = [b for b in bugs
                if b["priority"] == priority_filter and b["is_done"]]
    if not filtered:
        return None

    within = [b for b in filtered if b["bh"] <= sla_hours]
    pct    = round(len(within) / len(filtered) * 100)
    avg_bh = round(sum(b["bh"] for b in filtered) / len(filtered), 1)

    # In-flight (open) bugs of this priority
    open_bugs = [b for b in bugs
                 if b["priority"] == priority_filter and b["in_flight"]]
    breaching = [b for b in open_bugs if b["bh"] > sla_hours]

    return {
        "total":       len(filtered),
        "within_sla":  len(within),
        "pct":         pct,
        "on_target":   pct >= sla_pct_target,
        "avg_bh":      avg_bh,
        "open":        len(open_bugs),
        "breaching":   len(breaching),   # open bugs already past SLA
        "target_pct":  sla_pct_target,
        "target_bh":   sla_hours,
    }

# ── Main report ────────────────────────────────────────────────────────────────
def run_bug_sla():
    from datetime import date
    today = date.today()
    label = f"W/E {today.strftime('%b %d, %Y')} (last 12 weeks)"

    print(f"\n🐛 Bug SLA Report — {label}")
    print("=" * 60)
    print("   Urgent: ≤6 business hours (90% target)")
    print("   High:   ≤12 business hours / 1.5 days (75% target)")
    print("=" * 60)

    project_results = {}

    for project, jql_template in PROJECT_JQLS.items():
        jql = jql_template.format(days=LOOKBACK_DAYS)
        print(f"\n  Fetching {project} bugs...")
        issues = get_bugs(project, jql)
        bugs   = compute_bug_sla(issues)
        print(f"  → {len(issues)} bugs found ({sum(1 for b in bugs if b['is_done'])} resolved, "
              f"{sum(1 for b in bugs if b['in_flight'])} open)")

        urgent = sla_stats(bugs, "urgent", SLA_URGENT_BH, SLA_URGENT_PCT)
        high   = sla_stats(bugs, "high",   SLA_HIGH_BH,   SLA_HIGH_PCT)

        project_results[project] = {
            "urgent": urgent,
            "high":   high,
            "bugs":   bugs,
        }

        # Print summary
        for prio, stats, target_h in [("Urgent", urgent, SLA_URGENT_BH),
                                       ("High",   high,   SLA_HIGH_BH)]:
            if stats:
                ind = "✅" if stats["on_target"] else "⚠️"
                print(f"    {prio}: {stats['pct']}% within {target_h}h "
                      f"{ind} (goal: ≥{stats['target_pct']}%) "
                      f"— {stats['within_sla']}/{stats['total']} resolved "
                      f"| {stats['open']} open ({stats['breaching']} breaching)")
            else:
                print(f"    {prio}: No resolved bugs")

    return {"label": label, "projects": project_results}

# ── Slack message ──────────────────────────────────────────────────────────────
def build_slack_section(report):
    """
    Returns a Slack-formatted string to append to the cycle time message,
    or post standalone.
    """
    lines = [
        "*── Bug SLA — Urgent & High Priority ──*",
        f"_Labels: jira_escalated or kanban · last 12 weeks · "
        f"Business hours only (PT holidays excluded)_\n",
    ]

    for proj, data in report["projects"].items():
        urgent = data["urgent"]
        high   = data["high"]

        u_line = h_line = None

        if urgent:
            ind    = "✅" if urgent["on_target"] else "⚠️"
            breach = f" · {urgent['breaching']} open breaching SLA" if urgent["breaching"] else ""
            u_line = (f"Urgent (≤6h): *{urgent['pct']}%*{ind} "
                      f"_(goal: ≥{urgent['target_pct']}%)_ "
                      f"— {urgent['within_sla']}/{urgent['total']} resolved"
                      f"{breach}")
        else:
            u_line = "Urgent (≤6h): No resolved bugs"

        if high:
            ind    = "✅" if high["on_target"] else "⚠️"
            breach = f" · {high['breaching']} open breaching SLA" if high["breaching"] else ""
            h_line = (f"High (≤1.5d): *{high['pct']}%*{ind} "
                      f"_(goal: ≥{high['target_pct']}%)_ "
                      f"— {high['within_sla']}/{high['total']} resolved"
                      f"{breach}")
        else:
            h_line = "High (≤1.5d): No resolved bugs"

        lines += [
            f"*{proj}*",
            f"        {u_line}",
            f"        {h_line}",
            "",
        ]

    return "\n".join(lines)

def post_to_slack(report):
    if not SLACK_URL:
        print("\n⚠️  SLACK_WEBHOOK_URL not set — skipping Slack post")
        return

    msg = (
        f"*🐛 Bug SLA Report — {report['label']}*\n"
        f"_Urgent: 90% ≤6h · High: 75% ≤1.5 days · "
        f"Labels: jira_escalated or kanban_\n\n"
        + build_slack_section(report)
    )

    r = requests.post(SLACK_URL, json={"text": msg}, timeout=10)
    r.raise_for_status()
    print("\n✅ Posted to Slack #dtge_v2mom_sync")

# ── Entry point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    report = run_bug_sla()
    post_to_slack(report)
