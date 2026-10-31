"""
Bug SLA Report — CAECS, GTM, DTAL, DTK
=======================================
Measures resolution time for Urgent and High priority bugs
against SLA targets:
  - Urgent: 90% resolved within 6 business hours
  - High:   75% resolved within 1.5 business days (12 business hours)

Only counts bugs with label 'jira_escalated' OR 'kanban'.
Uses business hours only (Mon–Fri, excluding PT public holidays).
Blocked AND Test Blocked time is excluded from SLA timer (tracked separately).

Requirements:
    pip install requests python-dateutil openpyxl

Environment variables:
    JIRA_EMAIL         your Atlassian login email
    JIRA_API_TOKEN     Atlassian API token
    SLACK_WEBHOOK_URL  Slack webhook for #dtge_v2mom_sync (optional)
"""

import os
import json
import requests
import concurrent.futures
from datetime import datetime, date, timezone, timedelta
from dateutil import parser as dp

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

# ── Status classification ──────────────────────────────────────────────────────
# Same logic as cycle_time.py — Blocked/Test Blocked pause the SLA timer
BLOCKED_STATUSES = {"blocked", "test blocked", "release blocked"}
DONE_STATUSES    = {"done", "closed", "resolved", "ready for production"}

# ── Lookback window ────────────────────────────────────────────────────────────
LOOKBACK_DAYS = 84  # 12 weeks

# ── JQLs per project ──────────────────────────────────────────────────────────
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
    """Business hours Mon-Fri 09:00-17:00, PT holidays excluded."""
    if end_dt <= start_dt:
        return 0.0
    BH_START = 9
    BH_END   = 17
    total_bh = 0.0
    cursor   = datetime(start_dt.year, start_dt.month, start_dt.day, tzinfo=timezone.utc)
    while cursor < end_dt:
        if is_business_day(cursor):
            day_start = cursor.replace(hour=BH_START, minute=0, second=0, microsecond=0)
            day_end   = cursor.replace(hour=BH_END,   minute=0, second=0, microsecond=0)
            eff_start = max(day_start, start_dt)
            eff_end   = min(day_end,   end_dt)
            if eff_end > eff_start:
                total_bh += (eff_end - eff_start).total_seconds() / 3600
        cursor += timedelta(days=1)
    return total_bh

# ── Jira helpers ───────────────────────────────────────────────────────────────
def jira_get(path, params=None):
    r = requests.get(
        f"{JIRA_BASE}/{path}",
        params=params, auth=AUTH, headers=HEADERS, timeout=60
    )
    r.raise_for_status()
    return r.json()

def get_bugs(project, jql):
    issues, page_token = [], None
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
    return issues

def get_changelog(issue_key):
    """Fetch changelog for a bug to identify Blocked periods."""
    data = jira_get(f"issue/{issue_key}", {
        "expand": "changelog",
        "fields": "status",
    })
    return data.get("changelog", {}).get("histories", [])

# ── SLA Calculation with Blocked exclusion ─────────────────────────────────────
def compute_bug_bh(issue, histories):
    """
    Compute active business hours for a bug, excluding time in Blocked/Test Blocked.

    Timer runs from creation → resolution (or now if open).
    Pauses when entering Blocked/Test Blocked/Release Blocked.
    Resumes when leaving those states.
    """
    now     = datetime.now(timezone.utc)
    fields  = issue["fields"]
    status  = fields.get("status", {}).get("name", "").lower()
    created = dp.parse(fields["created"]).astimezone(timezone.utc)
    resolved = dp.parse(fields["resolutiondate"]).astimezone(timezone.utc) \
               if fields.get("resolutiondate") else None
    is_done  = status in DONE_STATUSES
    end_time = resolved if (is_done and resolved) else now

    # Build status transitions from changelog
    transitions = []
    for h in histories:
        for item in h.get("items", []):
            if item["field"] == "status":
                transitions.append({
                    "date": dp.parse(h["created"]).astimezone(timezone.utc),
                    "to":   (item.get("toString") or "").strip().lower(),
                    "from": (item.get("fromString") or "").strip().lower(),
                })
    transitions.sort(key=lambda x: x["date"])

    # Walk transitions to find blocked periods
    blocked_periods = []  # list of (start, end) tuples
    blocked_from    = None

    for t in transitions:
        if t["to"] in BLOCKED_STATUSES:
            if not blocked_from:
                blocked_from = t["date"]
        elif t["from"] in BLOCKED_STATUSES and blocked_from:
            blocked_periods.append((blocked_from, t["date"]))
            blocked_from = None

    # If still blocked at end_time
    if blocked_from:
        blocked_periods.append((blocked_from, end_time))

    # Total BH = BH(created → end_time) - BH spent in blocked periods
    total_bh   = business_hours_between(created, end_time)
    blocked_bh = sum(business_hours_between(bs, be) for bs, be in blocked_periods)
    active_bh  = max(0.0, total_bh - blocked_bh)

    return round(active_bh, 1), round(blocked_bh, 1), not is_done

# ── Process all bugs ───────────────────────────────────────────────────────────
def compute_bug_sla(issues, changelogs):
    now     = datetime.now(timezone.utc)
    results = []

    for issue in issues:
        fields   = issue["fields"]
        priority = (fields.get("priority") or {}).get("name", "Normal").lower()
        status   = fields.get("status", {}).get("name", "").lower()
        created  = dp.parse(fields["created"]).astimezone(timezone.utc)
        resolved = dp.parse(fields["resolutiondate"]).astimezone(timezone.utc) \
                   if fields.get("resolutiondate") else None
        is_done  = status in DONE_STATUSES

        histories  = changelogs.get(issue["key"], [])
        bh, blocked_bh, in_flight = compute_bug_bh(issue, histories)

        results.append({
            "key":        issue["key"],
            "summary":    fields.get("summary", ""),
            "priority":   priority,
            "status":     fields.get("status", {}).get("name", ""),
            "bh":         bh,
            "blocked_bh": blocked_bh,
            "is_done":    is_done,
            "in_flight":  in_flight,
            "created":    created.strftime("%Y-%m-%d"),
            "resolved":   resolved.strftime("%Y-%m-%d") if resolved else "open",
        })

    return results

def sla_stats(bugs, priority_filter, sla_hours, sla_pct_target):
    filtered = [b for b in bugs
                if b["priority"] == priority_filter and b["is_done"]]
    if not filtered:
        return None

    within = [b for b in filtered if b["bh"] <= sla_hours]
    pct    = round(len(within) / len(filtered) * 100)
    avg_bh = round(sum(b["bh"] for b in filtered) / len(filtered), 1)

    open_bugs = [b for b in bugs
                 if b["priority"] == priority_filter and b["in_flight"]]
    breaching = [b for b in open_bugs if b["bh"] > sla_hours]

    return {
        "total":      len(filtered),
        "within_sla": len(within),
        "pct":        pct,
        "on_target":  pct >= sla_pct_target,
        "avg_bh":     avg_bh,
        "open":       len(open_bugs),
        "breaching":  len(breaching),
        "target_pct": sla_pct_target,
        "target_bh":  sla_hours,
    }

# ── Main report ────────────────────────────────────────────────────────────────
def run_bug_sla():
    today = date.today()
    label = f"W/E {today.strftime('%b %d, %Y')} (last 12 weeks)"

    print(f"\n🐛 Bug SLA Report — {label}")
    print("=" * 60)
    print("   Urgent: ≤6 business hours (90% target)")
    print("   High:   ≤12 business hours / 1.5 days (75% target)")
    print("   Blocked/Test Blocked/Release Blocked time excluded from SLA timer")
    print("=" * 60)

    project_results = {}

    for project, jql_template in PROJECT_JQLS.items():
        jql    = jql_template.format(days=LOOKBACK_DAYS)
        print(f"\n  Fetching {project} bugs...")
        issues = get_bugs(project, jql)
        print(f"  → {len(issues)} bugs found — fetching changelogs...")

        # Fetch changelogs in parallel
        changelogs = {}
        with concurrent.futures.ThreadPoolExecutor(max_workers=10) as ex:
            futures = {ex.submit(get_changelog, i["key"]): i["key"] for i in issues}
            for future in concurrent.futures.as_completed(futures):
                key = futures[future]
                try:
                    changelogs[key] = future.result()
                except Exception as e:
                    print(f"    ⚠️  Could not fetch changelog for {key}: {e}")
                    changelogs[key] = []

        bugs   = compute_bug_sla(issues, changelogs)
        urgent = sla_stats(bugs, "urgent", SLA_URGENT_BH, SLA_URGENT_PCT)
        high   = sla_stats(bugs, "high",   SLA_HIGH_BH,   SLA_HIGH_PCT)

        project_results[project] = {
            "urgent": urgent,
            "high":   high,
            "bugs":   bugs,
        }

        resolved_count = sum(1 for b in bugs if b["is_done"])
        open_count     = sum(1 for b in bugs if b["in_flight"])
        print(f"  → {resolved_count} resolved, {open_count} open")

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

    # Save to JSON for cycle_time.py to pick up
    json_path = "bug_sla_results.json"
    export = {
        "label":    label,
        "projects": {
            proj: {
                "urgent": data["urgent"],
                "high":   data["high"],
                "bugs":   data["bugs"],
            }
            for proj, data in project_results.items()
        }
    }
    with open(json_path, "w") as f:
        json.dump(export, f, indent=2)
    print(f"\n📄 Results saved to {json_path}")

    return {"label": label, "projects": project_results}

# ── Slack message ──────────────────────────────────────────────────────────────
def build_slack_section(report):
    lines = [
        "*── Bug SLA — Urgent & High Priority ──*",
        f"_Labels: jira_escalated or kanban · last 12 weeks · "
        f"Business hours only · Blocked time excluded · PT holidays excluded_\n",
    ]

    for proj, data in report["projects"].items():
        urgent = data["urgent"]
        high   = data["high"]

        if urgent:
            ind    = "✅" if urgent["on_target"] else "⚠️"
            breach = f" · {urgent['breaching']} open breaching SLA" if urgent["breaching"] else ""
            u_line = (f"Urgent (≤6h): *{urgent['pct']}%*{ind} "
                      f"_(goal: ≥{urgent['target_pct']}%)_ "
                      f"— {urgent['within_sla']}/{urgent['total']} resolved{breach}")
        else:
            u_line = "Urgent (≤6h): No resolved bugs"

        if high:
            ind    = "✅" if high["on_target"] else "⚠️"
            breach = f" · {high['breaching']} open breaching SLA" if high["breaching"] else ""
            h_line = (f"High (≤1.5d): *{high['pct']}%*{ind} "
                      f"_(goal: ≥{high['target_pct']}%)_ "
                      f"— {high['within_sla']}/{high['total']} resolved{breach}")
        else:
            h_line = "High (≤1.5d): No resolved bugs"

        lines += [f"*{proj}*", f"        {u_line}", f"        {h_line}", ""]

    return "\n".join(lines)

def post_to_slack(report):
    if not SLACK_URL:
        print("\n⚠️  SLACK_WEBHOOK_URL not set — skipping Slack post")
        return

    msg = (
        f"*🐛 Bug SLA Report — {report['label']}*\n"
        f"_Urgent: 90% ≤6h · High: 75% ≤1.5 days · "
        f"Labels: jira_escalated or kanban · Blocked time excluded_\n\n"
        + build_slack_section(report)
    )

    r = requests.post(SLACK_URL, json={"text": msg}, timeout=10)
    r.raise_for_status()
    print("\n✅ Posted to Slack #dtge_v2mom_sync")

# ── Entry point ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    report = run_bug_sla()
    post_to_slack(report)
