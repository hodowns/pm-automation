#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
status_report.py
Weekly LCS Status Report generator — City of Houston (default PID F51119).
Trigger: "Status Report [PID]" in CLAUDE.md

Usage:
    py -3 -u -X utf8 status_report.py [PID] [--since YYYY-MM-DD] [--no-scan]
    py -3 -u -X utf8 status_report.py [PID] --push [--gh-org ORG] [--gh-repo REPO]
    py -3 -u -X utf8 status_report.py --gh-save-token --gh-token ghp_xxxx
"""

import argparse
import base64
import json
import os
import re
import subprocess
import sys
import urllib.request
import urllib.error
from datetime import datetime, timedelta
from pathlib import Path

# ═══════════════════════════════════════════════════════════════════
# CONFIG
# ═══════════════════════════════════════════════════════════════════

DEFAULT_PID = "F51119"
OUTPUT_DIR  = Path(r"C:\Users\hodowns\OneDrive - Cisco\Desktop\Status Reports")
WEBEX_TOKENS = Path(r"C:\Users\hodowns\webex-mcp-oauth\tokens.json")

GH_HOST        = "https://wwwin-github.cisco.com"
GH_API         = f"{GH_HOST}/api/v3"
GH_CONFIG_FILE = Path.home() / ".claude" / "sr_gh_config.json"
WEBEX_BASE     = "https://webexapis.com/v1"
TEMPLATE_FILE  = Path(__file__).parent / "status_report_template.html"

# ── Per-account lookup (add new accounts here) ─────────────────────
ACCOUNT_NAMES: dict[str, str] = {
    "F51119": "City of Houston",
    "F24373": "UChicago Medicine",
    "F48852": "MD-S-DHS",
    "F46267": "City of Houston – CURWB",
    "E55085": "Next Gen",
}

ACCOUNT_DOMAINS: dict[str, set] = {
    "F51119": {"houstontx.gov"},
    "F24373": {"uchicagomedicine.org", "meridianitinc.com"},
    "F48852": {"maryland.gov"},
    "F46267": {"houstontx.gov"},
}

WEBEX_ACCOUNT_KEYWORDS: dict[str, set] = {
    "F51119": {"f51119", "city of houston", "coh", "hits"},
    "F24373": {"f24373", "uchicago", "ucmc"},
    "F48852": {"f48852", "mddhs", "maryland"},
}

# Subject keywords used to match emails/Webex to a technology track (F51119-specific)
TRACK_KEYWORDS: dict[str, list] = {
    "curwb":           ["curwb", "ultra-reliable", "wifi for city events", "july 4"],
    "aci_vmm":         ["aci", "vmm", "apic", "fabric upgrade", "epg", "tenant migration"],
    "catalyst_center": ["catalyst center", "cat center", "dnac", "template", "health check"],
    "sdwan_support":   ["sdwan", "sd-wan", "vmanage", "dmvpn", "catalyst 8k", "cat8k"],
    "sdwan_as_code":   ["as code", "service as code", "terraform", "ansible", "gitlab", "git"],
    "aas":             ["architect as a service", "aas", "architecture advisory", "bgp"],
    "meraki":          ["meraki", "wireless", "wi-fi", "wifi", "ssid", "access point"],
    "cisco_iq":        ["cisco iq", "ciscoiq", "kpi", "psirt", "insights", "conformance"],
    "ise":             ["ise", "802.1x", "tacacs", "identity services", "zero trust"],
    "nexus":           ["nexus", "nxos", "nx-os"],
}

SHAREPOINT_BASE = Path(
    r"C:\Users\hodowns\Cisco\SLED United CX Delivery Team - CUSTOMER PROJECTS"
)

# ── Helpers ────────────────────────────────────────────────────────
def account_name(pid: str) -> str:
    return ACCOUNT_NAMES.get(pid, f"Account {pid}")

def account_domains(pid: str) -> set:
    return ACCOUNT_DOMAINS.get(pid, set())

def find_pid_folder(pid: str) -> Path | None:
    pid_up = pid.upper()
    if not SHAREPOINT_BASE.exists():
        return None
    for state_dir in SHAREPOINT_BASE.iterdir():
        if not state_dir.is_dir():
            continue
        for proj_dir in state_dir.iterdir():
            if proj_dir.is_dir() and pid_up in proj_dir.name.upper():
                return proj_dir
    return None

# ═══════════════════════════════════════════════════════════════════
# QUOTES (52 — rotates by ISO week)
# ═══════════════════════════════════════════════════════════════════

QUOTES = [
    {"text": "The impediment to action advances action. What stands in the way becomes the way.",
     "author": "Marcus Aurelius, Meditations"},
    {"text": "You have power over your mind, not outside events. Realize this, and you will find strength.",
     "author": "Marcus Aurelius, Meditations"},
    {"text": "Waste no more time arguing what a good man should be. Be one.",
     "author": "Marcus Aurelius, Meditations"},
    {"text": "When you arise in the morning, think of what a precious privilege it is to be alive — to breathe, to think, to enjoy, to love.",
     "author": "Marcus Aurelius, Meditations"},
    {"text": "It is not death that a man should fear, but he should fear never beginning to live.",
     "author": "Marcus Aurelius, Meditations"},
    {"text": "Accept the things to which fate binds you, and love the people with whom fate brings you together.",
     "author": "Marcus Aurelius, Meditations"},
    {"text": "The best revenge is to be unlike him who performed the injury.",
     "author": "Marcus Aurelius, Meditations"},
    {"text": "If it is not right, do not do it; if it is not true, do not say it.",
     "author": "Marcus Aurelius, Meditations"},
    {"text": "Very little is needed to make a happy life; it is all within yourself, in your way of thinking.",
     "author": "Marcus Aurelius, Meditations"},
    {"text": "Never let the future disturb you. You will meet it, if you have to, with the same weapons of reason which today arm you against the present.",
     "author": "Marcus Aurelius, Meditations"},
    {"text": "How much time he gains who does not look to see what his neighbor says or does or thinks.",
     "author": "Marcus Aurelius, Meditations"},
    {"text": "Nothing happens to any man that he is not formed by nature to bear.",
     "author": "Marcus Aurelius, Meditations"},
    {"text": "Our life is what our thoughts make it.",
     "author": "Marcus Aurelius, Meditations"},
    {"text": "Begin at once to live, and count each separate day as a separate life.",
     "author": "Seneca"},
    {"text": "He who is brave is free.",
     "author": "Seneca"},
    {"text": "Luck is what happens when preparation meets opportunity.",
     "author": "Seneca"},
    {"text": "We suffer more often in imagination than in reality.",
     "author": "Seneca"},
    {"text": "It is not the man who has too little, but the man who craves more, that is poor.",
     "author": "Seneca"},
    {"text": "No man was ever wise by chance.",
     "author": "Seneca"},
    {"text": "The first step is to establish that you are not bound to be what you were.",
     "author": "Seneca"},
    {"text": "Technology is best when it brings people together.",
     "author": "Matt Mullenweg"},
    {"text": "Innovation distinguishes between a leader and a follower.",
     "author": "Steve Jobs"},
    {"text": "The advance of technology is based on making it fit in so that you don't really even notice it.",
     "author": "Bill Gates"},
    {"text": "The true sign of intelligence is not knowledge but imagination.",
     "author": "Albert Einstein"},
    {"text": "First, solve the problem. Then, write the code.",
     "author": "John Johnson"},
    {"text": "Any sufficiently advanced technology is indistinguishable from magic.",
     "author": "Arthur C. Clarke"},
    {"text": "The internet is becoming the town square for the global village of tomorrow.",
     "author": "Bill Gates"},
    {"text": "The greatest glory in living lies not in never falling, but in rising every time we fall.",
     "author": "Nelson Mandela"},
    {"text": "In order to be irreplaceable, one must always be different.",
     "author": "Coco Chanel"},
    {"text": "The only way to do great work is to love what you do.",
     "author": "Steve Jobs"},
    {"text": "If you want to go fast, go alone. If you want to go far, go together.",
     "author": "African Proverb"},
    {"text": "Talent wins games, but teamwork and intelligence win championships.",
     "author": "Michael Jordan"},
    {"text": "Alone we can do so little; together we can do so much.",
     "author": "Helen Keller"},
    {"text": "Coming together is a beginning. Keeping together is progress. Working together is success.",
     "author": "Henry Ford"},
    {"text": "The strength of the team is each individual member. The strength of each member is the team.",
     "author": "Phil Jackson"},
    {"text": "The key is not the will to win. Everybody has that. It is the will to prepare to win.",
     "author": "Bobby Knight"},
    {"text": "Great things in business are never done by one person. They're done by a team of people.",
     "author": "Steve Jobs"},
    {"text": "Make every detail perfect and limit the number of details to perfect.",
     "author": "Jack Dorsey"},
    {"text": "Move fast and break things. Unless you are breaking stuff you are not moving fast enough.",
     "author": "Mark Zuckerberg"},
    {"text": "The best way to predict the future is to invent it.",
     "author": "Alan Kay"},
    {"text": "It does not matter how slowly you go as long as you do not stop.",
     "author": "Confucius"},
    {"text": "Start where you are. Use what you have. Do what you can.",
     "author": "Arthur Ashe"},
    {"text": "Success is not final, failure is not fatal: it is the courage to continue that counts.",
     "author": "Winston Churchill"},
    {"text": "You miss 100% of the shots you don't take.",
     "author": "Wayne Gretzky"},
    {"text": "Whether you think you can or you think you can't, you're right.",
     "author": "Henry Ford"},
    {"text": "I have not failed. I've just found 10,000 ways that won't work.",
     "author": "Thomas Edison"},
    {"text": "The way to get started is to quit talking and begin doing.",
     "author": "Walt Disney"},
    {"text": "Don't watch the clock; do what it does. Keep going.",
     "author": "Sam Levenson"},
    {"text": "Simplicity is the soul of efficiency.",
     "author": "Austin Freeman"},
    {"text": "Quality is not an act, it is a habit.",
     "author": "Aristotle"},
    {"text": "An investment in knowledge pays the best interest.",
     "author": "Benjamin Franklin"},
    {"text": "In the middle of difficulty lies opportunity.",
     "author": "Albert Einstein"},
]


# ═══════════════════════════════════════════════════════════════════
# INITIAL STATE (pre-populated from existing PPTX — F51119 Oct 2026)
# ═══════════════════════════════════════════════════════════════════

def _normalize_series_title(subj: str, pid: str, acct: str) -> str:
    """Strip PID, [EXTERNAL], account name from a meeting subject to get a clean series name."""
    t = subj
    t = re.sub(r'\[EXTERNAL\]|\[External\]|\[external\]', '', t, flags=re.IGNORECASE)
    t = re.sub(r'\bF\d{5}\b', '', t)
    t = re.sub(re.escape(acct), '', t, flags=re.IGNORECASE)
    t = re.sub(r'city of houston', '', t, flags=re.IGNORECASE)
    t = re.sub(r'\s{2,}', ' ', t)
    return t.strip(' -–—|:')


def auto_discover_tracks(pid: str) -> tuple[list[dict], list[dict]]:
    """Scan calendar for external meetings with PID.
    Returns (track_skeletons, raw_meetings) — both from a single calendar scan.
    """
    acct     = account_name(pid)
    meetings = scan_calendar(pid)   # raw meeting dicts (start, subject, duration, …)
    if not meetings:
        return [], []

    seen: dict[str, dict] = {}   # normalized_key → first meeting info
    for m in sorted(meetings, key=lambda x: x["start"]):
        title = _normalize_series_title(m["subject"], pid, acct)
        if not title:
            continue
        key = re.sub(r'\s+', ' ', title.lower())[:40]
        if key not in seen:
            seen[key] = {"title": title, "first_start": m["start"], "duration": m.get("duration", 60)}

    tracks = []
    for key, info in seen.items():
        safe_id = re.sub(r'[^a-z0-9]+', '_', key)[:24].strip('_') or f"track_{len(tracks)}"
        tracks.append({
            "id":               safe_id,
            "title":            info["title"],
            "pid":              pid,
            "date":             datetime.today().strftime("%B %d, %Y"),
            "coh_engineer":     "",
            "cisco_engineers":  "",
            "status":           "active",
            "project_start":    datetime.today().strftime("%m/%d/%y"),
            "project_end":      "",
            "budget_consumed":  0,
            "executive_summary": "",
            "key_activities":   [{"num": 1, "task": "", "status": "Not Started", "outcome": ""}],
            "project_status":   [],
            "next_steps":       [],
        })

    print(f"  Auto-discovered {len(tracks)} meeting series as slides:", flush=True)
    for t in tracks:
        print(f"    · {t['title']}", flush=True)
    return tracks, meetings


def build_generic_state(pid: str, tracks: list[dict]) -> dict:
    """Generic initial state for a new account, built from auto-discovered meeting series."""
    week  = datetime.today().isocalendar()[1]
    quote = QUOTES[week % len(QUOTES)]
    month = datetime.today().strftime("%B")
    acct  = account_name(pid)

    return {
        "_quotes": QUOTES,
        "meta": {
            "pid":       pid,
            "account":   acct,
            "generated": datetime.today().strftime("%Y-%m-%d"),
            "period":    f"Week of {datetime.today().strftime('%b %d, %Y')}",
        },
        "slide1": {
            "quote":  quote["text"],
            "author": quote["author"],
            "header": f"Happy {month}!",
        },
        "slide2": {
            "title": "TOP of MINDs",
            "items": [],
        },
        "slide3": {
            "title":            "Token / Credit Tracking",
            "total":            0,
            "time_consumed_pct": 0,
            "report_date":      datetime.today().strftime("%m/%d"),
            "tracks": [
                {"id": t["id"], "label": t["title"][:16], "consumed": 0, "total": 0}
                for t in tracks
            ],
        },
        "slide4": {
            "title":         "Workstreams & Initiatives",
            "total_credits": 0,
            "show_completed": False,
            "in_flight": [
                {"id": t["id"] + "_i", "name": t["title"], "progress": 0, "sub": ""}
                for t in tracks
            ],
            "completed": [],
        },
        "tracks":           tracks,
        "completed_tracks": [],
        "slide_cal": {
            "title":     "Meeting Calendar",
            "meetings":  [],
            "month":     datetime.today().strftime("%B %Y"),
            "year":      datetime.today().year,
            "month_num": datetime.today().month,
        },
    }


def build_initial_state(pid: str) -> dict:
    week = datetime.today().isocalendar()[1]
    quote = QUOTES[week % len(QUOTES)]
    today = datetime.today().strftime("%B %d, %Y")
    month = datetime.today().strftime("%B")

    return {
        "_quotes": QUOTES,
        "meta": {
            "pid": pid,
            "account": account_name(pid),
            "generated": datetime.today().strftime("%Y-%m-%d"),
            "period": f"Week of {datetime.today().strftime('%b %d, %Y')}",
        },
        "slide1": {
            "quote": quote["text"],
            "author": quote["author"],
            "header": f"Happy {month}!",
        },
        "slide2": {
            "title": "TOP of MINDs",
            "items": [
                {"text": "Technology roadmap discussion with CoH architects (planned for Nov during business review onsite visit – all 3 architects to be invited)", "done": False},
                {"text": "Contact Center migration resourcing in progress", "done": False},
                {"text": "Secure access migration SOW reviewed", "done": False},
                {"text": "CURWB deployment scope developed — downtown EDO corridor or 4th July event", "done": False},
                {"text": "HPD RMS implementation support — proactive TAC case needed", "done": False},
                {"text": "November QBR 1 planning — Week 2 of November TBD", "done": False},
                {"text": "Nexus DB Sessions with Charles (Talk about scheduling)", "done": False},
                {"text": "Calendar view of recurring external meetings — fused into status report", "done": True},
            ],
        },
        "slide_cal": {
            "title": "Meeting Calendar",
            "month": datetime.today().strftime("%B %Y"),
            "year": datetime.today().year,
            "month_num": datetime.today().month,
            "meetings": [],
        },
        "slide3": {
            "title": "Token Consumption Reporting",
            "total": 1500,
            "time_consumed_pct": 21,
            "report_date": datetime.today().strftime("%m/%d"),
            "tracks": [
                {"id": "aci_vmm",          "label": "ACI",             "consumed": 89,  "total": 200},
                {"id": "catalyst_center",  "label": "Catalyst Center", "consumed": 2,   "total": 200},
                {"id": "aas",              "label": "AaS",             "consumed": 2,   "total": 50},
                {"id": "sdwan_as_code",    "label": "As Code",         "consumed": 4,   "total": 250},
                {"id": "cisco_iq",         "label": "Cisco IQ",        "consumed": 2,   "total": 150},
                {"id": "ise",              "label": "ISE",             "consumed": 41,  "total": 250},
                {"id": "sdwan_support",    "label": "SDWAN",           "consumed": 14,  "total": 200},
                {"id": "meraki",           "label": "Meraki",          "consumed": 0,   "total": 200},
            ],
        },
        "slide4": {
            "title": "Workstreams & Initiatives",
            "total_credits": 1500,
            "show_completed": False,
            "in_flight": [
                {"id": "cc_lan",      "name": "Catalyst Center LAN standardization", "progress": 2, "sub": "Templating 101, Hands On Templating"},
                {"id": "aci_vmm_i",   "name": "ACI VMM Migration",                  "progress": 2, "sub": "VMM Migration: HITS, HPW cluster"},
                {"id": "sdwan_post",  "name": "SDWAN post-deployment support",       "progress": 2, "sub": "SDWAN Post Imp Support"},
                {"id": "automation",  "name": "Automation (as-code)",                "progress": 1, "sub": "SAC: SDWAN Service as Code production migration"},
                {"id": "aas_i",       "name": "Architecture as a Service",           "progress": 2, "sub": "AaS: Current Operations support and optimization planning advisory"},
                {"id": "iq_i",        "name": "Cisco IQ and Conformance",            "progress": 2, "sub": "IQ: Link, Integration to DNAC/SDWAN, Adoption"},
                {"id": "ise_i",       "name": "ISE Deployment",                      "progress": 1, "sub": "ISE: Securing SDWAN access, Healthcheck, 802.1x full network implementation"},
                {"id": "meraki_i",    "name": "Meraki Wireless",                     "progress": 2, "sub": "Ask The Expert"},
                {"id": "nexus_i",     "name": "Nexus",                               "progress": 0, "sub": "Coming soon"},
            ],
            "completed": [
                {"id": "aci_dep",     "name": "ACI Deployment",          "completed_date": "2026-09"},
                {"id": "cc_101",      "name": "Catalyst Center 101",     "completed_date": "2026-08"},
                {"id": "aci_tenant",  "name": "ACI Tenant Migration",    "completed_date": "2026-09-18"},
                {"id": "curwb_kt",    "name": "CURWB KT and 101",        "completed_date": "2026-10"},
            ],
        },
        "tracks": [
            {
                "id": "curwb",
                "title": "CURWB Deployment",
                "pid": "F46267",
                "date": today,
                "coh_engineer": "Jimbo Herrera",
                "cisco_engineers": "Roody Abmette, Sharath Kattemane Prabhakar",
                "status": "active",
                "project_start": "5/11/26",
                "project_end": "10/13/26",
                "budget_consumed": 77,
                "executive_summary": "Cisco is providing expert consulting services to deploy CURWB to support wireless networks for City of Houston events such as FIFA FanFest and July 4th. Request to consider another site by Omar F on 9/15 — to be discussed during scoping call.",
                "key_activities": [
                    {"num": 1, "task": "Meeting with CoH team on 9/9 to review scope for July 4th event", "status": "In Progress", "outcome": ""},
                    {"num": 2, "task": "Design and deployment build", "status": "Planned", "outcome": ""},
                    {"num": 3, "task": "Implementation execution & testing", "status": "Planned", "outcome": ""},
                    {"num": 4, "task": "Knowledge transfer session", "status": "Planned", "outcome": ""},
                ],
                "project_status": [
                    "Milestone 1: Assessment completed",
                    "Milestone 2: Solution Requirements Development completed",
                    "Milestone 3: Implementation Execution completed",
                    "Milestone 4: Knowledge Transfer session — PENDING",
                    "Additional support needed with July 4th event or potentially downtown EDO area — Cisco team to scope effort",
                ],
                "next_steps": [
                    "Coordinate scoping intake for July 4th event",
                    "Schedule KT sessions to provide overview of CURWB to HITS team for operational management",
                ],
            },
            {
                "id": "aci_vmm",
                "title": "ACI VMM Migration and Other",
                "pid": "F51119",
                "date": today,
                "coh_engineer": "Salim, Compute",
                "cisco_engineers": "Kennie Albritton",
                "status": "active",
                "project_start": "8/1/26",
                "project_end": "07/20/27",
                "budget_consumed": 68,
                "executive_summary": "Cisco is providing expert consulting services to improve operational efficiency, improve security posture, scalability, and adoption/value realization.",
                "key_activities": [
                    {"num": 1, "task": "VMM Migration: HITS Call Manager, HPW — 2/3", "status": "In Progress", "outcome": "90% Effort to Compute team, Kennie/Johnny M supporting"},
                    {"num": 2, "task": "HITS ACI Fabric Upgrade — VMM Dependent, Tgt 12/26", "status": "In Progress", "outcome": "Moving one cluster at a time"},
                    {"num": 3, "task": "HPD ACI Fabric Upgrade — Tgt 12/26", "status": "On Hold", "outcome": "ON HOLD — Recommending schedule after ACI Tenant migration, Tgt 11/05"},
                ],
                "project_status": [
                    "Working session commenced — Cisco General support underway",
                    "HITS Progress commenced 9/11",
                    "APIC Upgrades: Kennie on standby, TAC is first line of support",
                    "Progressed setup of ACI Simulator VM on ESXi/DVS to validate Ansible playbooks before production",
                ],
                "next_steps": [
                    "Continue with HPW teams to make progress on VMM Migration",
                    "Schedule ACI HPD Fabric Upgrade — Postponed until November",
                    "APIC Upgrade (TBD) — 4hr window, est 1.5hr",
                    "Switch Upgrade (TBD) — 8 hour window, est ~3-5hrs",
                ],
            },
            {
                "id": "catalyst_center",
                "title": "Catalyst Center Initiatives",
                "pid": "F51119",
                "date": today,
                "coh_engineer": "Chi, Eugene, Orestes, Wayne",
                "cisco_engineers": "Alec Franco",
                "status": "active",
                "project_start": "8/1/26",
                "project_end": "07/20/27",
                "budget_consumed": 45,
                "executive_summary": "Cisco is providing expert consulting services to adopt, expand, and improve Cat Center understanding.",
                "key_activities": [
                    {"num": 1, "task": "System Health check (30 min)", "status": "Completed", "outcome": "Operational Efficiency / Standardization"},
                    {"num": 2, "task": "Licensing Integration (SmartNet) (30 min)", "status": "Closed", "outcome": "Asset Optimization / Financial Governance"},
                    {"num": 3, "task": "Template Optimization — 5 of XX", "status": "In Progress", "outcome": "Standardization / Scalability"},
                    {"num": 4, "task": "Integration to Cisco IQ", "status": "Completed", "outcome": "Operational Efficiency / Proactive Insights"},
                    {"num": 5, "task": "RMA Workflow Automation", "status": "Not Started", "outcome": "Operational Efficiency / Reduced Downtime"},
                    {"num": 6, "task": "Upgrade Support & Use Cases (CoH)", "status": "Not Started", "outcome": "Adoption & Value Realization"},
                ],
                "project_status": [
                    "Health Check reviewed — Recommendations made",
                    "Re-enabling device controllability with a change request expected",
                    "Consolidation of virtual accounts cannot be done via Cat Center",
                    "HSEC License Application workflow reviewed",
                    "Integrated with Cisco IQ",
                    "Completing 1 template per session roughly (Slowburn — currently on step 7)",
                ],
                "next_steps": [
                    "Continue with Standard config template generation leveraging CoH's standard config sheet",
                    "Upon completion — review and condense",
                ],
            },
            {
                "id": "sdwan_support",
                "title": "SDWAN Support Initiatives",
                "pid": "F51119",
                "date": today,
                "coh_engineer": "Orestes, Eugene, Syed",
                "cisco_engineers": "Dan Beynon",
                "status": "active",
                "project_start": "8/1/26",
                "project_end": "07/20/27",
                "budget_consumed": 40,
                "executive_summary": "Cisco is providing expert consulting services to adopt, expand, and renew SDWAN infrastructure.",
                "key_activities": [
                    {"num": 1, "task": "DMVPN to SDWAN Pilot", "status": "Not Started", "outcome": "Standardization / Scalability"},
                    {"num": 2, "task": "Integration to Cisco ISE - TACACs", "status": "Not Started", "outcome": "Operational Efficiency / Reduced Downtime"},
                    {"num": 3, "task": "Integration to Cisco IQ", "status": "In Progress", "outcome": "Operational Efficiency / Proactive Insights"},
                    {"num": 4, "task": "Continued General Support", "status": "Ongoing", "outcome": "Standardization / Scalability"},
                ],
                "project_status": [
                    "Validated Cisco IQ",
                    "Reviewed Station 17 Failover",
                    "SDWAN SR205159088: Subnet Interface Design Initiative deep dive",
                    "Meraki to SDWAN: Routing issue discovered and remedied",
                    "vManage Security Remediation: Established upgrade path to 20.15.6.1",
                    "Cloud ops process for cloud-hosted vManage upgrades shared",
                    "SDWAN Bandwidth report shared",
                ],
                "next_steps": [
                    "Track 1109 TAC Case & Re-run Cat8K Upgrades — Submit crash files to TAC, isolate branch instability",
                    "Complete Case w/TAC — Monitor for resolution",
                    "vManage disk utilization data and sizing recommendations sent",
                ],
            },
            {
                "id": "sdwan_as_code",
                "title": "SDWAN Service as Code",
                "pid": "F51119",
                "date": today,
                "coh_engineer": "HITS",
                "cisco_engineers": "Dan Carr, Omar Saffouri",
                "status": "active",
                "project_start": "8/1/26",
                "project_end": "07/20/27",
                "budget_consumed": 21,
                "executive_summary": "Cisco is providing expert consulting services to adopt, expand, and renew SDWAN via infrastructure as code.",
                "key_activities": [
                    {"num": 1, "task": "Lab Build (Validation)", "status": "Completed", "outcome": "Risk Reduction / Design Validation"},
                    {"num": 2, "task": "Knowledge Transfer", "status": "In Progress", "outcome": "Accelerated Value / Time-to-Deploy"},
                    {"num": 3, "task": "Migration to Production", "status": "In Progress", "outcome": "Accelerated Value / Time-to-Deploy"},
                    {"num": 4, "task": "Service as Code (Automation)", "status": "Not Started", "outcome": "Operational Efficiency / Standardization"},
                ],
                "project_status": [
                    "Re-initiation decision point",
                    "Lab Validation complete",
                    "CAF Direction selected",
                    "Aided in Git and VS Code acquisition",
                    "Set up feature branch in GitLab",
                    "Designed tailored curriculum covering Terraform mechanics and CI/CD pipelines",
                    "Developing step-by-step onboarding walkthrough guides for City network engineers",
                ],
                "next_steps": [
                    "Continue to verify git access",
                    "Familiarization with git commands",
                    "Leverage Onboarding Guide part 3 (in CoH SharePoint)",
                    "Talk through moving to production",
                ],
            },
            {
                "id": "aas",
                "title": "Architect as a Service",
                "pid": "F51119",
                "date": today,
                "coh_engineer": "HITS",
                "cisco_engineers": "Nizami, Nicole, Antonio",
                "status": "active",
                "project_start": "8/1/26",
                "project_end": "07/20/27",
                "budget_consumed": 21,
                "executive_summary": "Cisco is providing expert consulting services to aid in BGP Plan, Design and broader architecture advisory.",
                "key_activities": [
                    {"num": 1, "task": "Umbrella at HPL", "status": "In Progress", "outcome": "New RR in for Umbrella direct passive support"},
                    {"num": 2, "task": "HPL, HEC, DB Link Design and implementation", "status": "In Progress", "outcome": ""},
                    {"num": 3, "task": "DMVPN to SDWAN HLD", "status": "In Progress", "outcome": "Build requirements and publish HLD outlining eventual 55-site migration"},
                    {"num": 4, "task": "Error Disabled on Ports deep dive", "status": "Planned", "outcome": "SR 701503702 case opened by Orestes"},
                ],
                "project_status": [
                    "Evaluated infrastructure connectivity dependencies — HPL-to-Databank link redundancy",
                    "FTD firewall NAT hairpinning reviewed",
                    "Alignment with Netsync on security/VRF architectures",
                ],
                "next_steps": [
                    "Q/A Email to Omar F on DMVPN Gap questions",
                    "CoH Questionnaire released for discussion and completion",
                    "Update on ISPs and HPL internet redundancy?",
                ],
            },
            {
                "id": "meraki",
                "title": "Wireless / Meraki Support",
                "pid": "F51119",
                "date": today,
                "coh_engineer": "Chi Vu / Mustafa / Eugene",
                "cisco_engineers": "Abhishek Mehrotra",
                "status": "active",
                "project_start": "8/1/26",
                "project_end": "07/20/27",
                "budget_consumed": 25,
                "executive_summary": "Cisco is providing expert consulting services to improve current facets of the network via working sessions. Documentation is not desired — SME sessions for remediation support to address city-wide wireless issues.",
                "key_activities": [
                    {"num": 1, "task": "Ask The Expert Sessions", "status": "Not Started", "outcome": "Optimization"},
                    {"num": 2, "task": "(NEW) Wireless Issue at Remote site", "status": "Not Started", "outcome": "Monitoring as agenda topic"},
                ],
                "project_status": [
                    "Pending meeting series",
                    "MTU settings review — cases #14509991 and #14523097 opened with Meraki requesting TCP MSS Clamping (1400)",
                    "That did fix all intermittent issues with CoH_Guest SSID",
                    "Read access to DB granted",
                    "Artifact sharing complete",
                ],
                "next_steps": [
                    "1333 N. Freeway HT site best practices and review ongoing",
                    "Continue to review Meraki DB and supply best practice thoughts",
                    "Chi Vu — Monitoring update: roaming/disconnections improved after disabling client balancing?",
                    "Abhishek Mehrotra — Heat map + survey file review: AP placement confirmed? Final RF parameter recommendations",
                    "Chi Vu — Maintenance window status: full test profile changes applied?",
                    "Eugene / Chi — Device 2 follow-up: any recurrence of authentication failure or offline behavior?",
                ],
            },
            {
                "id": "cisco_iq",
                "title": "Cisco IQ",
                "pid": "F51119",
                "date": today,
                "coh_engineer": "HITS",
                "cisco_engineers": "Ankur Ruhal",
                "status": "active",
                "project_start": "8/1/26",
                "project_end": "07/20/27",
                "budget_consumed": 15,
                "executive_summary": "CiscoIQ provides assets and security insights all within one portal. Objective to leverage IQ for actionable insights in replacement of traditional KPI reporting. New process for monthly review of IQ insights for software conformance/PSIRTs/EoL planning.",
                "key_activities": [
                    {"num": 1, "task": "Demo IQ", "status": "Completed", "outcome": "Operational Efficiency"},
                    {"num": 2, "task": "Onboarding/Integrations", "status": "Completed", "outcome": "Integration of data connectors"},
                    {"num": 3, "task": "Data validation and transition from KPI reporting", "status": "In Progress", "outcome": "Leverage Ankur to transition from traditional KPI reporting to IQ"},
                    {"num": 4, "task": "Insights Review (Monthly)", "status": "In Progress", "outcome": "Ankur to lead KPI review sessions using IQ portal"},
                    {"num": 5, "task": "Reporting Development/Automation", "status": "Not Started", "outcome": "Operational Efficiency"},
                ],
                "project_status": [
                    "Reinitiation of conversation to onboard (Ankur)",
                    "Link established",
                    "CSPC: INTEGRATED",
                    "Catalyst Center: INTEGRATED",
                    "SDWAN Manager: INTEGRATED",
                    "Webex: INTEGRATED",
                    "Intersight: INTEGRATED",
                    "PSIRT Review for NX devices",
                ],
                "next_steps": [
                    "Attain config info in IQ for ACI devices (In Progress)",
                ],
            },
            {
                "id": "ise",
                "title": "ISE Initiatives — On Hold Until Post Pilot",
                "pid": "F51119",
                "date": today,
                "coh_engineer": "Chi Vu",
                "cisco_engineers": "Antonio Brown",
                "status": "active",
                "project_start": "8/1/26",
                "project_end": "07/20/27",
                "budget_consumed": 21,
                "executive_summary": "Cisco is providing expert consulting services to improve operational efficiency, improve security posture, scalability, and adoption/value realization via ISE.",
                "key_activities": [
                    {"num": 1, "task": "Securing SDWAN Authentication", "status": "Not Started", "outcome": "Security Posture / Zero Trust"},
                    {"num": 2, "task": "Health-Check (8th & 9th Floor)", "status": "Not Started", "outcome": "Operational Efficiency"},
                    {"num": 3, "task": "Design Review (Deployment Planning)", "status": "Not Started", "outcome": "Scalability / Accelerated Value"},
                    {"num": 4, "task": "Other Use Case Support (CoH)", "status": "Not Started", "outcome": "Adoption & Value Realization"},
                    {"num": 5, "task": "802.1x Support", "status": "In Progress", "outcome": "Reduce Risk / Operational Efficiency"},
                ],
                "project_status": [
                    "Antonio & Chi 30 min one-off to discuss nuances in 802.1X",
                ],
                "next_steps": [
                    "Offer support should Chi need it post Antonio email write-up",
                ],
            },
            {
                "id": "nexus",
                "title": "Nexus — Coming Soon",
                "pid": "F51119",
                "date": today,
                "coh_engineer": "Charles",
                "cisco_engineers": "TBD",
                "status": "active",
                "project_start": "TBD",
                "project_end": "07/20/27",
                "budget_consumed": 0,
                "executive_summary": "Placeholder — Nexus DB sessions and NX-OS support scope to be defined.",
                "key_activities": [
                    {"num": 1, "task": "Scope definition meeting with Charles", "status": "Planned", "outcome": ""},
                ],
                "project_status": [
                    "Scope TBD — sessions with Charles to be scheduled",
                ],
                "next_steps": [
                    "Talk about scheduling Nexus DB sessions with Charles",
                ],
            },
        ],
        "completed_tracks": [
            {
                "id": "aci_tenant",
                "title": "ACI EPG Tenant Migration",
                "pid": "F51119",
                "completed_date": "2026-09-18",
                "coh_engineer": "Salim",
                "cisco_engineers": "Kennie Albritton",
                "key_activities": [
                    {"num": 1, "task": "Preparation: Capture, Backup, Map, automate", "status": "Completed", "outcome": "Old EPGs rooted out 7/29"},
                    {"num": 2, "task": "Setup: Connectivity check, Evac host, port reconfig, MAC verification", "status": "Completed", "outcome": ""},
                    {"num": 3, "task": "Execution: Monitor, Disable old routing, Enable new routing, Policy Check, Verify", "status": "Completed", "outcome": "Conducted Go-no-go 8/26, 9/1, 9/17"},
                    {"num": 4, "task": "Finalize: Post Checks", "status": "Completed", "outcome": ""},
                ],
                "project_status": ["Closed 9/22 — COMPLETED 9/18"],
                "next_steps": [],
            },
        ],
    }


# ═══════════════════════════════════════════════════════════════════
# WEBEX HELPERS
# ═══════════════════════════════════════════════════════════════════

def _load_webex_token() -> str:
    try:
        return json.loads(WEBEX_TOKENS.read_text())["access_token"]
    except Exception:
        return ""


def _webex_get(path: str, params: dict = None):
    import urllib.request
    token = _load_webex_token()
    if not token:
        return None
    url = f"{WEBEX_BASE}/{path.lstrip('/')}"
    if params:
        from urllib.parse import urlencode
        url += "?" + urlencode(params)
    req = urllib.request.Request(url, headers={
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
    })
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return json.loads(resp.read())
    except Exception as e:
        status = getattr(e, 'code', None)
        if status == 401:
            print("  Warning: Webex token expired (401). Skipping Webex scan.", flush=True)
        else:
            print(f"  Warning: Webex error — {e}", flush=True)
        return None


# ═══════════════════════════════════════════════════════════════════
# OUTLOOK SCANNING
# ═══════════════════════════════════════════════════════════════════

def scan_outlook(since: datetime, pid: str = DEFAULT_PID) -> list[dict]:
    """Return list of {subject, body, sender, date, track_id} for account emails since cutoff."""
    domains = account_domains(pid)
    try:
        import win32com.client
    except ImportError:
        print("  Warning: pywin32 not available — skipping Outlook scan.", flush=True)
        return []
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns      = outlook.GetNamespace("MAPI")
    except Exception as e:
        print(f"  Warning: Could not connect to Outlook — {e}", flush=True)
        return []

    cutoff_str = since.strftime("%m/%d/%Y %I:%M %p")
    emails = []
    seen = set()

    def try_item(item):
        try:
            if item.Class != 43:
                return
            eid = item.EntryID
            if eid in seen:
                return
            seen.add(eid)
            sender  = (item.SenderEmailAddress or "").lower()
            subject = item.Subject or ""
            body    = (item.Body or "")[:1500]
            # Check if CoH domain involved
            rec_domains = set()
            try:
                for r in item.Recipients:
                    addr = (r.Address or "").lower()
                    if "@" in addr:
                        rec_domains.add(addr.split("@")[1])
            except Exception:
                pass
            if not domains or not (any(d in sender for d in domains) or any(d in rd for d in domains for rd in rec_domains)):
                return
            track_id = _detect_track(subject + " " + body)
            emails.append({
                "subject": subject,
                "body":    body[:800],
                "sender":  sender,
                "date":    str(item.ReceivedTime)[:10],
                "track_id": track_id,
            })
        except Exception:
            pass

    try:
        for folder_id in (6, 5):
            folder = ns.GetDefaultFolder(folder_id)
            restriction = f"[ReceivedTime] >= '{cutoff_str}'"
            for item in folder.Items.Restrict(restriction):
                try_item(item)
    except Exception as e:
        print(f"  Warning: Outlook scan error — {e}", flush=True)

    print(f"  Outlook: found {len(emails)} CoH emails since {since.strftime('%Y-%m-%d')}", flush=True)
    return emails


def _detect_track(text: str) -> str:
    t = text.lower()
    for track_id, keywords in TRACK_KEYWORDS.items():
        if any(kw in t for kw in keywords):
            return track_id
    return "general"


# ═══════════════════════════════════════════════════════════════════
# WEBEX SCANNING
# ═══════════════════════════════════════════════════════════════════

def scan_webex(since: datetime) -> list[dict]:
    """Return list of {text, sender, date, track_id} from CoH Webex spaces."""
    rooms_data = _webex_get("rooms", {"type": "group", "max": 200, "sortBy": "lastactivity"})
    if not rooms_data:
        print("  Webex: skipped (no token or error)", flush=True)
        return []

    matching_rooms = [
        r for r in rooms_data.get("items", [])
        if any(kw in r.get("title", "").lower() for kw in WEBEX_KEYWORDS)
    ]
    print(f"  Webex: found {len(matching_rooms)} matching spaces", flush=True)

    messages = []
    since_str = since.strftime("%Y-%m-%dT%H:%M:%S.000Z")
    for room in matching_rooms[:5]:
        data = _webex_get("messages", {"roomId": room["id"], "after": since_str, "max": 50})
        if not data:
            continue
        for msg in data.get("items", []):
            txt  = (msg.get("text") or msg.get("markdown") or "")[:800]
            if not txt.strip():
                continue
            track_id = _detect_track(txt)
            messages.append({
                "text":     txt,
                "sender":   msg.get("personEmail", ""),
                "date":     msg.get("created", "")[:10],
                "room":     room["title"],
                "track_id": track_id,
            })

    print(f"  Webex: collected {len(messages)} messages", flush=True)
    return messages


# ═══════════════════════════════════════════════════════════════════
# CALENDAR SCANNING
# ═══════════════════════════════════════════════════════════════════

def scan_calendar(pid: str) -> list[dict]:
    """Return external Outlook calendar meetings for current month where subject contains PID."""
    try:
        import win32com.client
    except ImportError:
        print("  Warning: pywin32 not available — skipping calendar scan.", flush=True)
        return []
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        ns  = outlook.GetNamespace("MAPI")
        cal = ns.GetDefaultFolder(9)
        items = cal.Items
        items.IncludeRecurrences = True
        items.Sort("[Start]")
    except Exception as e:
        print(f"  Warning: Could not access Outlook calendar — {e}", flush=True)
        return []

    today = datetime.today()
    yr, mo = today.year, today.month
    if mo == 12:
        last_day = 31
    else:
        last_day = (datetime(yr, mo + 1, 1) - timedelta(days=1)).day
    start_str = f"{mo:02d}/01/{yr} 00:00 AM"
    end_str   = f"{mo:02d}/{last_day:02d}/{yr} 11:59 PM"

    try:
        filtered = items.Restrict(f"[Start] >= '{start_str}' AND [Start] <= '{end_str}'")
    except Exception as e:
        print(f"  Warning: Calendar restriction failed — {e}", flush=True)
        return []

    EXCLUDE_RE = re.compile(
        r"send end of day recap|send_eod|eod recap|timecards due|hd travel|hd not taking",
        re.IGNORECASE,
    )
    results = []
    seen_keys = set()          # deduplicate recurring occurrences on the same day
    for item in filtered:
        try:
            subj = item.Subject or ""
            # Skip programmatic calendar reminders
            if EXCLUDE_RE.search(subj):
                continue
            # Must contain the exact PID
            if pid not in subj:
                continue
            # Must be flagged as external
            if "[external]" not in subj.lower():
                continue
            start_iso = item.Start.strftime("%Y-%m-%dT%H:%M") if hasattr(item.Start, "strftime") else str(item.Start)[:16]
            end_iso   = item.End.strftime("%Y-%m-%dT%H:%M")   if hasattr(item.End,   "strftime") else str(item.End)[:16]
            # Deduplicate: same subject on same day (recurring) → keep one
            dedup_key = (subj.strip().lower(), start_iso[:10])
            if dedup_key in seen_keys:
                continue
            seen_keys.add(dedup_key)
            try:    duration = item.Duration
            except: duration = 60
            try:    required = item.RequiredAttendees or ""
            except: required = ""
            try:    location = item.Location or ""
            except: location = ""
            try:    body = (item.Body or "")[:200].strip().replace("\r\n", " ").replace("\n", " ")
            except: body = ""
            results.append({
                "subject":  subj,
                "start":    start_iso,
                "end":      end_iso,
                "duration": duration,
                "required": required,
                "location": location,
                "body":     body,
            })
        except Exception:
            pass

    print(f"  Calendar: found {len(results)} external {pid} meetings for {today.strftime('%B %Y')}", flush=True)
    return results


# ═══════════════════════════════════════════════════════════════════
# ENRICH STATE WITH EMAIL/WEBEX CONTENT
# ═══════════════════════════════════════════════════════════════════

def enrich_state(state: dict, emails: list[dict], webex_msgs: list[dict]) -> dict:
    """Append new email/Webex summaries as project_status bullets on relevant tracks."""
    from collections import defaultdict
    by_track: dict = defaultdict(list)

    for e in emails:
        if e["track_id"] and e["track_id"] != "general":
            by_track[e["track_id"]].append(f"[Email {e['date']}] {e['subject']}")

    for m in webex_msgs:
        if m["track_id"] and m["track_id"] != "general":
            snippet = m["text"][:120].replace("\n", " ")
            by_track[m["track_id"]].append(f"[Webex {m['date']}] {snippet}")

    for track in state.get("tracks", []):
        tid = track["id"]
        if tid in by_track:
            existing = set(track.get("project_status", []))
            for bullet in by_track[tid]:
                if bullet not in existing:
                    track["project_status"].insert(0, bullet)

    return state


# ═══════════════════════════════════════════════════════════════════
# HTML GENERATION
# ═══════════════════════════════════════════════════════════════════

def generate_report(state: dict, output_path: Path):
    if not TEMPLATE_FILE.exists():
        print(f"ERROR: Template not found at {TEMPLATE_FILE}", flush=True)
        sys.exit(1)
    template = TEMPLATE_FILE.read_text(encoding="utf-8")
    html = template.replace("__REPORT_DATA__", json.dumps(state, ensure_ascii=False))
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html, encoding="utf-8")
    print(f"  Report saved: {output_path}", flush=True)


# ═══════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════

# ═══════════════════════════════════════════════════════════════════
# MEETING NOTES SCANNING
# ═══════════════════════════════════════════════════════════════════

def scan_meeting_notes(pid: str, since: datetime) -> list[dict]:
    """Read .md meeting note files from the SharePoint project folder since cutoff."""
    pid_folder = find_pid_folder(pid)
    if not pid_folder:
        print(f"  Meeting notes: SharePoint folder not found for {pid}", flush=True)
        return []

    # Look for meeting notes in common subfolder names
    note_dirs = []
    for candidate in ["Meeting Minutes", "Meeting Notes", "Notes", "Minutes"]:
        d = pid_folder / candidate
        if d.exists():
            note_dirs.append(d)
    if not note_dirs:
        note_dirs = [pid_folder]   # fall back to project root

    notes = []
    for note_dir in note_dirs:
        for md_file in sorted(note_dir.glob("**/*.md"), key=lambda f: f.stat().st_mtime, reverse=True):
            try:
                mtime = datetime.fromtimestamp(md_file.stat().st_mtime)
                if mtime < since:
                    continue
                content = md_file.read_text(encoding="utf-8", errors="ignore")
                # Parse frontmatter date if present
                date_match = re.search(r'^date:\s*(\d{4}-\d{2}-\d{2})', content, re.MULTILINE)
                note_date = date_match.group(1) if date_match else mtime.strftime("%Y-%m-%d")
                notes.append({
                    "filename": md_file.name,
                    "date":     note_date,
                    "content":  content[:6000],   # cap per note to manage token budget
                })
            except Exception:
                pass

    print(f"  Meeting notes: found {len(notes)} notes since {since.strftime('%Y-%m-%d')}", flush=True)
    return notes


# ═══════════════════════════════════════════════════════════════════
# AI SLIDE ENRICHMENT
# ═══════════════════════════════════════════════════════════════════

def ai_enrich_slides(state: dict, notes: list[dict], emails: list[dict]) -> dict:
    """Use Claude to read meeting notes and emails, then populate project_status
    and next_steps on each track/slide with real content."""
    if not notes and not emails:
        return state

    try:
        import anthropic
    except ImportError:
        print("  AI enrichment skipped — anthropic not installed.", flush=True)
        return state

    tracks = state.get("tracks", [])
    if not tracks:
        return state

    # Build a compact notes block (most recent first, up to 12k chars total)
    notes_block = ""
    budget = 12000
    for n in notes:
        chunk = f"\n\n=== MEETING NOTE: {n['filename']} ({n['date']}) ===\n{n['content']}"
        if len(notes_block) + len(chunk) > budget:
            break
        notes_block += chunk

    # Build a compact emails block (subject + first 300 chars of body)
    email_lines = []
    for e in emails[:15]:
        email_lines.append(
            f"- [{e['date']}] {e['subject']}\n  {(e.get('body',''))[:300].replace(chr(10),' ')}"
        )
    emails_block = "\n".join(email_lines)

    track_list = "\n".join(f"- {t['title']}" for t in tracks)

    prompt = f"""You are updating a weekly project status report for {state.get('meta',{}).get('account','the account')} ({state.get('meta',{}).get('pid','')}).

The report has one slide per meeting series or technology track. Here are the slide titles:
{track_list}

Below are the meeting notes and emails from the past week.

MEETING NOTES:
{notes_block}

RECENT EMAILS:
{emails_block}

For each slide title above, extract content from the notes/emails and return a JSON object.
Keys are slide titles (match exactly). Values are objects with:
  "project_status": list of 2-5 concise bullet strings (what happened, what was decided, current state)
  "next_steps": list of 1-4 specific action items (who does what by when, if mentioned)

Rules:
- Only include a slide if there is relevant content in the notes/emails for it
- Be specific — use actual details, names, decisions, not generic filler
- project_status bullets describe the current state and what was accomplished
- next_steps are forward-looking actions
- If an action item has an owner named in the notes, include their name
- Keep each bullet to one sentence

Return ONLY valid JSON, no explanation."""

    try:
        client = anthropic.Anthropic()
        response = client.messages.create(
            model="claude-sonnet-4-5-20250929",
            max_tokens=3000,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = response.content[0].text.strip()
        # Strip markdown code fences if present
        raw = re.sub(r'^```(?:json)?\s*', '', raw)
        raw = re.sub(r'\s*```$', '', raw)
        enriched: dict = json.loads(raw)
    except Exception as e:
        print(f"  AI enrichment failed: {e}", flush=True)
        return state

    today_str = datetime.today().strftime("%m/%d")

    def _text(b):
        return b.get("text", "") if isinstance(b, dict) else b

    def _is_raw_ref(b):
        t = _text(b)
        return t.startswith("[Email ") or t.startswith("[Webex ")

    def _wrap(items):
        return [{"text": s, "added": today_str} if isinstance(s, str) else s for s in items]

    updated = 0
    for track in tracks:
        data = enriched.get(track["title"], {})
        if not data:
            continue
        if data.get("project_status"):
            existing = [b for b in track.get("project_status", []) if not _is_raw_ref(b)]
            track["project_status"] = _wrap(data["project_status"]) + existing
        if data.get("next_steps"):
            existing_ns = [b for b in track.get("next_steps", []) if not _is_raw_ref(b)]
            track["next_steps"] = _wrap(data["next_steps"]) + existing_ns
        updated += 1

    print(f"  AI enrichment: updated {updated}/{len(tracks)} slides from meeting notes", flush=True)
    return state


# ═══════════════════════════════════════════════════════════════════
# GITHUB PAGES COLLABORATION
# ═══════════════════════════════════════════════════════════════════

def gh_get_token(arg_token: str = "") -> str:
    """Return GitHub PAT from --gh-token arg, GH_PAT env var, or stored config."""
    if arg_token:
        return arg_token
    if os.environ.get("GH_PAT"):
        return os.environ["GH_PAT"]
    if GH_CONFIG_FILE.exists():
        try:
            return json.loads(GH_CONFIG_FILE.read_text())["token"]
        except Exception:
            pass
    return ""


def gh_save_token(token: str):
    """Persist a GitHub PAT to the config file for future use."""
    GH_CONFIG_FILE.parent.mkdir(parents=True, exist_ok=True)
    GH_CONFIG_FILE.write_text(json.dumps({"token": token}))
    print(f"  GitHub token saved to {GH_CONFIG_FILE}", flush=True)


def _gh_api(url: str, token: str, data: dict | None = None, method: str = "GET"):
    """Simple GitHub API call using stdlib urllib. Returns (body_dict, status_code)."""
    req = urllib.request.Request(url, method=method)
    req.add_header("Authorization", f"token {token}")
    req.add_header("Accept", "application/vnd.github.v3+json")
    if data:
        req.add_header("Content-Type", "application/json")
        req.data = json.dumps(data).encode("utf-8")
    try:
        with urllib.request.urlopen(req, timeout=20) as r:
            return json.loads(r.read()), r.status
    except urllib.error.HTTPError as e:
        body = {}
        try:
            body = json.loads(e.read())
        except Exception:
            pass
        return body, e.code
    except Exception as e:
        print(f"    GitHub request error: {e}", flush=True)
        return {}, 0


def pull_state_from_github(org: str, repo: str, branch: str, token: str, pid: str) -> dict | None:
    """Fetch the current state.json from GitHub and return it, or None on failure."""
    if not token:
        return None
    file_path = f"data/{pid}_state.json"
    url       = f"{GH_API}/repos/{org}/{repo}/contents/{file_path}?ref={branch}"
    result, code = _gh_api(url, token)
    if code != 200:
        return None
    try:
        raw = base64.b64decode(result["content"].replace("\n", "")).decode("utf-8")
        return json.loads(raw)
    except Exception:
        return None


def push_state_to_github(state: dict, org: str, repo: str, branch: str, token: str) -> bool:
    """Commit F51119_state.json into data/ on the given GitHub branch."""
    if not token:
        print("  GitHub push skipped — no token. Use --gh-token or run --gh-save-token.", flush=True)
        return False

    pid       = state.get("meta", {}).get("pid", DEFAULT_PID)
    file_path = f"data/{pid}_state.json"
    url       = f"{GH_API}/repos/{org}/{repo}/contents/{file_path}"

    # Fetch current SHA (required to update an existing file)
    existing, code = _gh_api(url, token)
    sha = existing.get("sha") if code == 200 else None

    content = base64.b64encode(
        json.dumps(state, indent=2, ensure_ascii=False).encode("utf-8")
    ).decode()

    payload: dict = {
        "message": f"Status Report {datetime.today().strftime('%Y-%m-%d')} — auto update",
        "content": content,
        "branch":  branch,
    }
    if sha:
        payload["sha"] = sha

    result, code = _gh_api(url, token, data=payload, method="PUT")
    if code in (200, 201):
        print(f"  GitHub: pushed {file_path} → {org}/{repo}@{branch}", flush=True)
        return True
    else:
        msg = result.get("message", "unknown error")
        print(f"  GitHub push failed ({code}): {msg}", flush=True)
        return False


def push_index_to_github(state: dict, org: str, repo: str, branch: str, token: str) -> bool:
    """Push the GH-pages-ready index.html (done once to set up Pages, or to update template)."""
    if not TEMPLATE_FILE.exists():
        print(f"  Template not found: {TEMPLATE_FILE}", flush=True)
        return False

    template = TEMPLATE_FILE.read_text(encoding="utf-8")

    # Inject real values so the page knows which repo to talk to
    html = (template
        .replace("__REPORT_DATA__", json.dumps(state, ensure_ascii=False))
        .replace("'__GH_ORG__'",    f"'{org}'")
        .replace("'__GH_REPO__'",   f"'{repo}'")
        .replace("'__GH_BRANCH__'", f"'{branch}'")
    )

    url      = f"{GH_API}/repos/{org}/{repo}/contents/index.html"
    existing, code = _gh_api(url, token)
    sha = existing.get("sha") if code == 200 else None

    payload: dict = {
        "message": "Deploy LCS Status Report page",
        "content": base64.b64encode(html.encode("utf-8")).decode(),
        "branch":  branch,
    }
    if sha:
        payload["sha"] = sha

    result, code = _gh_api(url, token, data=payload, method="PUT")
    if code in (200, 201):
        print(f"  GitHub: pushed index.html → {org}/{repo}@{branch}", flush=True)
        pages_url = f"{GH_HOST}/pages/{org}/{repo}/"
        print(f"  Live URL (once Pages is enabled): {pages_url}", flush=True)
        return True
    else:
        msg = result.get("message", "unknown error")
        print(f"  GitHub index.html push failed ({code}): {msg}", flush=True)
        return False


def _migrate_bullet_dates(state: dict) -> None:
    """Ensure all project_status and next_steps items are {text, added} objects.
    Items that are still plain strings get tagged with the report date so the
    item-date feature works correctly on every run."""
    today = datetime.today().strftime("%m/%d")
    for track in state.get("tracks", []) + state.get("completed_tracks", []):
        for field in ("project_status", "next_steps"):
            items = track.get(field)
            if not isinstance(items, list):
                continue
            for i, item in enumerate(items):
                if isinstance(item, str):
                    items[i] = {"text": item, "added": today}


def main():
    parser = argparse.ArgumentParser(description="LCS Status Report Generator")
    parser.add_argument("pid",             nargs="?", default=DEFAULT_PID, help="Project ID (default F51119)")
    parser.add_argument("--since",         default=None,            help="Scan start date YYYY-MM-DD (default: 7 days ago)")
    parser.add_argument("--no-scan",       action="store_true",     help="Skip email/Webex scanning")
    # GitHub Pages collaboration
    parser.add_argument("--push",          action="store_true",     help="Push state.json to GitHub Pages repo after generating")
    parser.add_argument("--deploy-pages",  action="store_true",     help="Also push index.html (run once when setting up Pages)")
    parser.add_argument("--gh-org",        default="cx-usps-auto",  help="GitHub org  (default: cx-usps-auto)")
    parser.add_argument("--gh-repo",       default="status-report-presentation",help="GitHub repo (default: status-report-presentation)")
    parser.add_argument("--gh-branch",     default="main",          help="GitHub branch (default: main)")
    parser.add_argument("--gh-token",      default="",              help="GitHub PAT (or set GH_PAT env var, or use --gh-save-token)")
    parser.add_argument("--gh-save-token", action="store_true",     help="Save --gh-token to ~/.claude/sr_gh_config.json for future use")
    args = parser.parse_args()

    # One-time token save
    if args.gh_save_token:
        if not args.gh_token:
            print("ERROR: pass --gh-token ghp_xxxx together with --gh-save-token", flush=True)
            sys.exit(1)
        gh_save_token(args.gh_token)
        if not args.push and not args.deploy_pages:
            print("Token saved. Run with --push to deploy.", flush=True)
            return

    pid  = args.pid.upper()
    since = (
        datetime.strptime(args.since, "%Y-%m-%d")
        if args.since
        else datetime.today() - timedelta(days=7)
    )

    print(f"\nStatus Report — {pid} / {account_name(pid)}", flush=True)
    print(f"Scanning from: {since.strftime('%Y-%m-%d')}", flush=True)

    state_file = OUTPUT_DIR / f"{pid}_state.json"

    # When pushing to GitHub, pull the live GitHub state first so any edits
    # the manager made in the browser are preserved as the base for this run.
    if args.push or args.deploy_pages:
        gh_token_early = gh_get_token(getattr(args, 'gh_token', ''))
        if gh_token_early:
            gh_state = pull_state_from_github(args.gh_org, args.gh_repo, args.gh_branch, gh_token_early, pid)
            if gh_state:
                print("  Pulled live state from GitHub (browser edits preserved)", flush=True)
                state_file.write_text(json.dumps(gh_state, indent=2, ensure_ascii=False), encoding="utf-8")

    # Load previous state or build fresh
    if state_file.exists():
        try:
            state = json.loads(state_file.read_text(encoding="utf-8"))
            # Refresh meta and quote for this week
            state["meta"]["generated"] = datetime.today().strftime("%Y-%m-%d")
            state["meta"]["period"]    = f"Week of {datetime.today().strftime('%b %d, %Y')}"
            week = datetime.today().isocalendar()[1]
            state["slide1"]["quote"]  = QUOTES[week % len(QUOTES)]["text"]
            state["slide1"]["author"] = QUOTES[week % len(QUOTES)]["author"]
            month = datetime.today().strftime("%B")
            state["slide1"]["header"] = f"Happy {month}!"
            _migrate_bullet_dates(state)
            print("  Loaded previous state (manual edits preserved)", flush=True)
        except Exception as e:
            print(f"  Warning: could not load state — {e}. Using fresh data.", flush=True)
            state = build_initial_state(pid) if pid == DEFAULT_PID else build_generic_state(pid, [])
    else:
        if pid == DEFAULT_PID:
            state = build_initial_state(pid)
            print("  Building fresh state from PPTX baseline (F51119 / City of Houston)", flush=True)
        else:
            # New account — auto-discover meeting series from calendar
            print(f"\n  New account detected: {account_name(pid)} ({pid})", flush=True)
            print("  Scanning calendar to discover meeting series...", flush=True)
            discovered, cal_meetings = auto_discover_tracks(pid)
            state = build_generic_state(pid, discovered)
            # Populate calendar slide from the same scan — no second calendar call needed
            state["slide_cal"]["meetings"] = cal_meetings
            if not discovered:
                print("  No external meetings found on calendar. Starting with blank slides.", flush=True)
                print("  Use '+ Track' in the nav bar to add meeting slides manually.", flush=True)
            print(f"  First run complete — edit slides in the browser, then Export State to preserve changes.", flush=True)

    # Ensure slide_cal exists in loaded state
    if "slide_cal" not in state:
        state["slide_cal"] = {"title": "Meeting Calendar", "meetings": []}
    state["slide_cal"]["month"]     = datetime.today().strftime("%B %Y")
    state["slide_cal"]["year"]      = datetime.today().year
    state["slide_cal"]["month_num"] = datetime.today().month

    # Scan for new content
    emails, webex_msgs = [], []
    if not args.no_scan:
        print("\nScanning Outlook...", flush=True)
        emails = scan_outlook(since, pid)

        print("Scanning Webex...", flush=True)
        webex_msgs = scan_webex(since)

        print("Scanning meeting notes...", flush=True)
        notes = scan_meeting_notes(pid, since)

        print("Scanning calendar...", flush=True)
        state["slide_cal"]["meetings"] = scan_calendar(pid)

        # AI enrichment from meeting notes + emails (replaces subject-line bullets)
        if notes or emails:
            print("Running AI enrichment from meeting notes...", flush=True)
            state = ai_enrich_slides(state, notes, emails)
        elif webex_msgs:
            # Fallback: use Webex snippets as before
            state = enrich_state(state, [], webex_msgs)

    # Save state for next run (preserves manual edits + new scan content)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    state_file.write_text(json.dumps(state, indent=2, ensure_ascii=False), encoding="utf-8")

    # Generate HTML
    today_str = datetime.today().strftime("%Y-%m-%d")
    out_file  = OUTPUT_DIR / f"{pid}_status_{today_str}.html"
    print("\nGenerating report...", flush=True)
    generate_report(state, out_file)

    # Open in browser
    subprocess.Popen(["start", "", str(out_file)], shell=True)
    print(f"\nDone. Report opened in browser.", flush=True)
    print(f"State saved to: {state_file}", flush=True)
    print(f"Tip: Use 'Export State' in the report to save manual edits for next week.", flush=True)

    # GitHub Pages push (--push flag)
    if args.push or args.deploy_pages:
        gh_token = gh_get_token(args.gh_token)
        if not gh_token:
            print("\nGitHub push skipped — no token found.", flush=True)
            print("Run:  py -3 status_report.py --gh-save-token --gh-token ghp_YOUR_TOKEN", flush=True)
        else:
            print(f"\nPushing to GitHub ({args.gh_org}/{args.gh_repo})...", flush=True)
            push_state_to_github(state, args.gh_org, args.gh_repo, args.gh_branch, gh_token)
            if args.deploy_pages:
                push_index_to_github(state, args.gh_org, args.gh_repo, args.gh_branch, gh_token)
            elif args.push:
                print("  Tip: add --deploy-pages on first run to also push index.html.", flush=True)


if __name__ == "__main__":
    main()
