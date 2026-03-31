#!/usr/bin/env python3
"""
Outlook Mailagent v4
– Prioritering & kategorisering
– Selvlæring (tracker svarmønstre og justerer prioritet automatisk)
– Afsender-læring
– Opgaveliste
– Opfølgningspåminder
– Mødeforberedelse (Outlook Kalender)
– Teams-integration
– Rapport direkte på mail
"""

import os, json, re, sys, requests
from datetime import datetime, timedelta, timezone
from pathlib import Path
from dotenv import load_dotenv

# ── Konfiguration ─────────────────────────────────────────────────────────────
load_dotenv(Path(__file__).parent / ".env")

AZURE_CLIENT_ID   = os.getenv("AZURE_CLIENT_ID", "")
AZURE_TENANT_ID   = os.getenv("AZURE_TENANT_ID", "common")
ANTHROPIC_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
HOURS_BACK        = int(os.getenv("HOURS_BACK", "4"))
SEND_EMAIL_REPORT = os.getenv("SEND_EMAIL_REPORT", "true").lower() == "true"
OPFØLGNING_TIMER  = int(os.getenv("OPFØLGNING_TIMER", "24"))

BASE             = Path(__file__).parent
REPORTS_DIR      = BASE / "rapporter"
TOKEN_CACHE_FILE = BASE / ".token_cache.json"
OPGAVER_JSON     = BASE / "opgaver.json"
OPGAVER_MD       = BASE / "OPGAVER.md"
AFSENDERE_FILE   = BASE / "vigtige_afsendere.json"
AFVENTER_FILE    = BASE / "afventer_svar.json"
LAERING_FILE     = BASE / "laering.json"

SCOPES = [
    "https://graph.microsoft.com/Mail.Read",
    "https://graph.microsoft.com/Mail.Send",
    "https://graph.microsoft.com/User.Read",
    "https://graph.microsoft.com/Calendars.Read",
    "https://graph.microsoft.com/Chat.Read",
    "https://graph.microsoft.com/ChannelMessage.Read.All",
]


# ── Hjælpefunktioner ──────────────────────────────────────────────────────────
def load_vigtige_afsendere() -> list:
    if AFSENDERE_FILE.exists():
        return [a.lower().strip() for a in json.loads(AFSENDERE_FILE.read_text()).get("afsendere", [])]
    return []


# ── Selvlæring ────────────────────────────────────────────────────────────────
def load_laering() -> dict:
    """Indlæs læringsdata – scorer pr. afsender baseret på svarmønstre."""
    if LAERING_FILE.exists():
        return json.loads(LAERING_FILE.read_text(encoding="utf-8"))
    return {"afsendere": {}, "sidst_opdateret": None}

def gem_laering(data: dict):
    LAERING_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")

def opdater_laering(indkomne: list, sendte: list):
    """
    Sammenlign indkomne mails med sendte svar.
    Øg score for afsendere du svarer hurtigt på – sænk for dem du aldrig svarer på.
    """
    data = load_laering()
    scorer = data.get("afsendere", {})
    nu = datetime.now(timezone.utc)

    # Lav et lookup: conversationId → sendt tidspunkt
    sendte_conv = {}
    for s in sendte:
        conv = s.get("conversationId")
        sent_tid = s.get("sentDateTime", "")
        if conv and sent_tid:
            sendte_conv[conv] = sent_tid

    for mail in indkomne:
        afsender = mail.get("from", {}).get("emailAddress", {}).get("address", "").lower()
        if not afsender:
            continue

        conv = mail.get("conversationId")
        modtaget_str = mail.get("receivedDateTime", "")

        if afsender not in scorer:
            scorer[afsender] = {
                "navn": mail.get("from", {}).get("emailAddress", {}).get("name", afsender),
                "score": 50,        # Starter neutralt på 50
                "hurtigt_svar": 0,  # Antal gange svaret inden 2 timer
                "aldrig_svar": 0,   # Antal gange ikke svaret
                "total": 0,
            }

        scorer[afsender]["total"] += 1

        if conv and conv in sendte_conv:
            # Beregn svartid
            try:
                modtaget = datetime.fromisoformat(modtaget_str.replace("Z", "+00:00"))
                sendt    = datetime.fromisoformat(sendte_conv[conv].replace("Z", "+00:00"))
                svartid  = (sendt - modtaget).total_seconds() / 3600  # timer

                if svartid <= 2:
                    # Svarede inden for 2 timer – meget vigtig
                    scorer[afsender]["score"] = min(100, scorer[afsender]["score"] + 5)
                    scorer[afsender]["hurtigt_svar"] += 1
                elif svartid <= 24:
                    # Svarede inden for en dag – moderat vigtig
                    scorer[afsender]["score"] = min(100, scorer[afsender]["score"] + 2)
            except Exception:
                pass
        else:
            # Ingen svar fundet – sænk score en smule
            scorer[afsender]["score"] = max(0, scorer[afsender]["score"] - 1)
            scorer[afsender]["aldrig_svar"] += 1

    data["afsendere"] = scorer
    data["sidst_opdateret"] = nu.isoformat()
    gem_laering(data)
    return scorer

def get_laerte_vigtige(scorer: dict, grænse: int = 75) -> list:
    """Returner afsendere med score over grænsen – automatisk lært."""
    return [
        email for email, info in scorer.items()
        if info.get("score", 0) >= grænse
    ]

def generer_laering_oversigt(scorer: dict) -> str:
    """Lav en kort tekst til rapporten om hvad agenten har lært."""
    if not scorer:
        return ""
    top = sorted(scorer.items(), key=lambda x: x[1].get("score", 0), reverse=True)[:5]
    linjer = []
    for email, info in top:
        score = info.get("score", 0)
        navn  = info.get("navn", email)
        emoji = "🔴" if score >= 75 else "🟡" if score >= 50 else "🟢"
        linjer.append(f"{emoji} {navn} (score: {score})")
    return " · ".join(linjer)

def load_opgaver() -> list:
    return json.loads(OPGAVER_JSON.read_text(encoding="utf-8")) if OPGAVER_JSON.exists() else []

def save_opgaver(opgaver: list):
    OPGAVER_JSON.write_text(json.dumps(opgaver, ensure_ascii=False, indent=2), encoding="utf-8")

def tilføj_opgaver(nye: list) -> list:
    eks = load_opgaver()
    eks_tekster = {o.get("opgave", "") for o in eks}
    for o in nye:
        if o.get("opgave", "") not in eks_tekster:
            eks.append(o)
    save_opgaver(eks)
    return eks

def generer_opgaver_md(alle: list):
    åbne = [o for o in alle if not o.get("udført")]
    lukkede = [o for o in alle if o.get("udført")]
    linjer = [
        "# 📋 Opgaveliste – Outlook Mailagent\n\n",
        f"_Opdateret: {datetime.now().strftime('%d/%m/%Y %H:%M')}_\n\n",
        f"**{len(åbne)} åbne** · {len(lukkede)} udførte\n\n---\n\n## ✅ Åbne opgaver\n\n",
    ]
    for o in åbne:
        frist = f" _(frist: {o['frist']})_" if o.get("frist") else ""
        linjer += [f"- [ ] **{o.get('opgave','')}**{frist}  \n",
                   f"  _Fra: {o.get('afsender','')} · {o.get('dato','')} · {o.get('mail_emne','')}_\n\n"]
    if not åbne:
        linjer.append("_Ingen åbne opgaver_ 🎉\n\n")
    OPGAVER_MD.write_text("".join(linjer), encoding="utf-8")

def load_afventer() -> list:
    return json.loads(AFVENTER_FILE.read_text(encoding="utf-8")) if AFVENTER_FILE.exists() else []

def save_afventer(afventer: list):
    AFVENTER_FILE.write_text(json.dumps(afventer, ensure_ascii=False, indent=2), encoding="utf-8")


# ── Microsoft Graph – godkendelse ─────────────────────────────────────────────
def get_access_token() -> str:
    try:
        from msal import PublicClientApplication, SerializableTokenCache
    except ImportError:
        print("❌  pip3 install msal"); sys.exit(1)

    cache = SerializableTokenCache()
    if TOKEN_CACHE_FILE.exists():
        cache.deserialize(TOKEN_CACHE_FILE.read_text())

    app = PublicClientApplication(
        AZURE_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{AZURE_TENANT_ID}",
        token_cache=cache,
    )
    accounts = app.get_accounts()
    result = app.acquire_token_silent(SCOPES, account=accounts[0]) if accounts else None

    if not result:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            raise Exception("Login-flow fejl: " + str(flow))
        print(f"\n{'─'*60}\n🔐  LOG IND:\n    {flow['message']}\n{'─'*60}\n")
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" not in result:
        raise Exception("Login fejlede: " + result.get("error_description", "ukendt"))

    TOKEN_CACHE_FILE.write_text(cache.serialize())
    return result["access_token"]


# ── Microsoft Graph – data ────────────────────────────────────────────────────
def fetch_user_info(token: str) -> dict:
    r = requests.get("https://graph.microsoft.com/v1.0/me",
                     headers={"Authorization": f"Bearer {token}"}, timeout=15)
    r.raise_for_status(); return r.json()

def fetch_emails(token: str, hours_back: int = 4) -> list:
    since = (datetime.now(timezone.utc) - timedelta(hours=hours_back)).strftime("%Y-%m-%dT%H:%M:%SZ")
    url = (
        "https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages"
        f"?$filter=receivedDateTime ge {since}"
        "&$select=id,subject,from,receivedDateTime,bodyPreview,importance,isRead,hasAttachments,conversationId"
        "&$orderby=receivedDateTime desc&$top=100"
    )
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    r.raise_for_status(); return r.json().get("value", [])

def fetch_sent_emails(token: str, hours_back: int = 4) -> list:
    """Hent sendte mails for at tjekke om bruger har svaret."""
    since = (datetime.now(timezone.utc) - timedelta(hours=hours_back)).strftime("%Y-%m-%dT%H:%M:%SZ")
    url = (
        "https://graph.microsoft.com/v1.0/me/mailFolders/SentItems/messages"
        f"?$filter=sentDateTime ge {since}"
        "&$select=conversationId,subject,sentDateTime&$top=50"
    )
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=30)
    r.raise_for_status(); return r.json().get("value", [])

def fetch_calendar_events(token: str) -> list:
    """Hent dagens og morgendagens møder fra Outlook-kalenderen."""
    now   = datetime.now(timezone.utc)
    end   = now + timedelta(hours=24)
    start_str = now.strftime("%Y-%m-%dT%H:%M:%SZ")
    end_str   = end.strftime("%Y-%m-%dT%H:%M:%SZ")
    url = (
        f"https://graph.microsoft.com/v1.0/me/calendarview"
        f"?startDateTime={start_str}&endDateTime={end_str}"
        "&$select=subject,start,end,attendees,bodyPreview,organizer,isAllDay"
        "&$orderby=start/dateTime&$top=20"
    )
    try:
        r = requests.get(url, headers={"Authorization": f"Bearer {token}"}, timeout=20)
        r.raise_for_status(); return r.json().get("value", [])
    except Exception as e:
        print(f"    ⚠️  Kalender ikke tilgængelig: {e}"); return []

def fetch_teams_messages(token: str) -> list:
    """Hent seneste Teams-beskeder fra alle chats."""
    beskeder = []
    try:
        # Hent chats
        r = requests.get(
            "https://graph.microsoft.com/v1.0/me/chats?$select=id,topic,chatType&$top=10",
            headers={"Authorization": f"Bearer {token}"}, timeout=20
        )
        r.raise_for_status()
        chats = r.json().get("value", [])

        since = (datetime.now(timezone.utc) - timedelta(hours=4)).strftime("%Y-%m-%dT%H:%M:%SZ")

        for chat in chats[:5]:  # Max 5 chats
            try:
                r2 = requests.get(
                    f"https://graph.microsoft.com/v1.0/me/chats/{chat['id']}/messages"
                    f"?$filter=createdDateTime ge {since}&$top=10",
                    headers={"Authorization": f"Bearer {token}"}, timeout=15
                )
                r2.raise_for_status()
                msgs = r2.json().get("value", [])
                for m in msgs:
                    if m.get("messageType") == "message" and m.get("body", {}).get("content"):
                        indhold = re.sub(r'<[^>]+>', '', m["body"]["content"])[:200]
                        if indhold.strip():
                            beskeder.append({
                                "chat": chat.get("topic") or "Direkte besked",
                                "afsender": m.get("from", {}).get("user", {}).get("displayName", "?"),
                                "indhold": indhold.strip(),
                                "tidspunkt": m.get("createdDateTime", ""),
                            })
            except Exception:
                continue
    except Exception as e:
        print(f"    ⚠️  Teams ikke tilgængelig: {e}")
    return beskeder

def send_report_email(token: str, user: dict, html: str, subject: str):
    r = requests.post(
        "https://graph.microsoft.com/v1.0/me/sendMail",
        headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"},
        json={"message": {
            "subject": subject,
            "body": {"contentType": "HTML", "content": html},
            "toRecipients": [{"emailAddress": {"address": user.get("mail", "")}}],
        }, "saveToSentItems": False},
        timeout=30,
    )
    r.raise_for_status()


# ── Opfølgningspåminder ───────────────────────────────────────────────────────
def opdater_opfølgning(emails: list, sent_emails: list, analyse: dict) -> tuple:
    """
    Registrer høj-prioritets mails som afventer svar.
    Returner (overskredet, afventer) – lister af mails.
    """
    afventer = load_afventer()
    nu = datetime.now(timezone.utc)

    # Email-ID'er på sendte mails (conversation-match)
    sendte_conv = {m.get("conversationId") for m in sent_emails if m.get("conversationId")}

    # Fjern mails vi nu har svaret på
    afventer = [a for a in afventer if a.get("conversationId") not in sendte_conv]

    # Tilføj nye høj-prioritets mails
    afventer_conv = {a.get("conversationId") for a in afventer}
    høj_numre = {h.get("nummer") for h in analyse.get("høj_prioritet", [])}

    for i, mail in enumerate(emails, 1):
        if i in høj_numre and mail.get("conversationId") not in afventer_conv:
            afventer.append({
                "id": mail.get("id"),
                "conversationId": mail.get("conversationId"),
                "emne": mail.get("subject", ""),
                "afsender": mail.get("from", {}).get("emailAddress", {}).get("name", "?"),
                "modtaget": mail.get("receivedDateTime", ""),
                "registreret": nu.isoformat(),
            })

    # Find overskredet mails (ingen svar efter OPFØLGNING_TIMER timer)
    grænse = nu - timedelta(hours=OPFØLGNING_TIMER)
    overskredet = []
    stadig_afventer = []
    for a in afventer:
        reg = datetime.fromisoformat(a["registreret"].replace("Z", "+00:00"))
        if reg < grænse:
            overskredet.append(a)
        else:
            stadig_afventer.append(a)

    save_afventer(stadig_afventer + overskredet)
    return overskredet, stadig_afventer


# ── Mødeforberedelse ──────────────────────────────────────────────────────────
def forbered_møder(møder: list, emails: list) -> list:
    """Lav en kort briefing for hvert møde baseret på relaterede mails."""
    if not møder:
        return []

    forberedelser = []
    for møde in møder:
        # Find deltagere
        deltagere = [
            a.get("emailAddress", {}).get("address", "").lower()
            for a in møde.get("attendees", [])
            if a.get("type") != "resource"
        ]

        # Find relevante mails fra deltagerne
        relaterede = []
        for mail in emails:
            afsender_mail = mail.get("from", {}).get("emailAddress", {}).get("address", "").lower()
            emne = mail.get("subject", "").lower()
            møde_emne = møde.get("subject", "").lower()

            # Match på afsender eller emne-lighed
            if afsender_mail in deltagere or any(ord in emne for ord in møde_emne.split() if len(ord) > 4):
                relaterede.append(f"- {mail.get('subject','')} (fra {mail.get('from',{}).get('emailAddress',{}).get('name','?')}): {mail.get('bodyPreview','')[:100]}")

        # Tidspunkt
        start_raw = møde.get("start", {}).get("dateTime", "")
        try:
            start_dt = datetime.fromisoformat(start_raw.replace("Z", "+00:00"))
            tidspunkt = start_dt.strftime("%H:%M")
        except Exception:
            tidspunkt = start_raw[:5] if start_raw else "?"

        end_raw = møde.get("end", {}).get("dateTime", "")
        try:
            end_dt = datetime.fromisoformat(end_raw.replace("Z", "+00:00"))
            slut = end_dt.strftime("%H:%M")
        except Exception:
            slut = "?"

        forberedelser.append({
            "emne": møde.get("subject", "Møde"),
            "tidspunkt": f"{tidspunkt}–{slut}",
            "deltagere": len(deltagere),
            "relaterede_mails": relaterede[:3],
            "all_day": møde.get("isAllDay", False),
        })

    return forberedelser


# ── AI-analyse ────────────────────────────────────────────────────────────────
def analyze_with_claude(emails: list, vigtige_afsendere: list,
                        teams_msgs: list, møder: list) -> dict:
    try:
        import anthropic
    except ImportError:
        print("❌  pip3 install anthropic"); sys.exit(1)

    client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)

    afs_sektion = ""
    if vigtige_afsendere:
        afs_sektion = f"\nVIGTIGE AFSENDERE (altid høj prioritet): {', '.join(vigtige_afsendere)}\n"

    møde_sektion = ""
    if møder:
        møde_liste = "\n".join(
            f"- {m.get('subject','?')} kl. {m.get('start',{}).get('dateTime','?')[:16]}"
            for m in møder[:5]
        )
        møde_sektion = f"\nDAGENS MØDER (tag hensyn til disse ved prioritering):\n{møde_liste}\n"

    teams_sektion = ""
    if teams_msgs:
        teams_liste = "\n".join(
            f"- [{m['chat']}] {m['afsender']}: {m['indhold'][:150]}"
            for m in teams_msgs[:10]
        )
        teams_sektion = f"\nTEAMS-BESKEDER (seneste 4 timer):\n{teams_liste}\n"

    mail_linjer = []
    for i, m in enumerate(emails, 1):
        s = m.get("from", {}).get("emailAddress", {})
        mail_linjer.append(
            f"[Mail {i}]\nAfsender: {s.get('name','?')} <{s.get('address','')}>\n"
            f"Emne: {m.get('subject','(intet)')}\nTidspunkt: {m.get('receivedDateTime','')}\n"
            f"Forhåndsvisning: {m.get('bodyPreview','')[:300]}\n"
            f"Ulæst: {'Ja' if not m.get('isRead') else 'Nej'} | "
            f"Vedhæftning: {'Ja' if m.get('hasAttachments') else 'Nej'}\n"
        )

    prompt = f"""Du er en professionel mailagent. Analyser disse {len(emails)} mails.
{afs_sektion}{møde_sektion}{teams_sektion}

MAILS:
{"---".join(mail_linjer)}

Returner KUN gyldigt JSON:
{{
  "oversigt": "2-3 sætninger om indbakkens tilstand",
  "statistik": {{"total": {len(emails)}, "ulæste": 0, "høj": 0, "medium": 0, "lav": 0}},
  "høj_prioritet": [
    {{"nummer": 1, "emne": "...", "afsender": "navn <email>", "handling": "konkret handling", "deadline": null, "tidspunkt": "HH:MM"}}
  ],
  "medium_prioritet": [...],
  "lav_prioritet": [...],
  "deadlines": [
    {{"mail_nummer": 1, "emne": "...", "afsender": "...", "deadline": "beskrivelse", "dage_tilbage": null}}
  ],
  "opgaver": [
    {{"mail_nummer": 1, "emne": "...", "afsender": "navn <email>", "opgave": "konkret opgave", "frist": null}}
  ],
  "teams_vigtige": [
    {{"chat": "...", "afsender": "...", "besked": "kort sammendrag", "kræver_handling": true}}
  ]
}}

REGLER:
- HØJ: Kræver svar i dag, vigtig afsender, deadline <48t, juridisk/finansielt
- MEDIUM: Besvar inden 1-3 dage
- LAV: Nyhedsbreve, notifikationer, FYI
- Tag hensyn til møder: mails relateret til dagens møder løftes i prioritet
- Svar KUN på dansk – KUN JSON"""

    msg = client.messages.create(
        model="claude-opus-4-6", max_tokens=4000,
        messages=[{"role": "user", "content": prompt}]
    )
    raw = msg.content[0].text.strip()
    match = re.search(r"\{.*\}", raw, re.DOTALL)
    if not match:
        raise ValueError("Ugyldig JSON: " + raw[:300])
    return json.loads(match.group())


# ── HTML-rapport ──────────────────────────────────────────────────────────────
def _kort(item: dict, farve: str, ikon: str) -> str:
    dl = f'<span class="dl-badge">⏰ {item["deadline"]}</span>' if item.get("deadline") else ""
    return f"""<div class="card" style="border-left:4px solid {farve}">
      <div class="ch"><span>{ikon}</span>
        <div class="cm"><div class="cs">{item.get('emne','')}</div>
        <div class="ca">fra {item.get('afsender','')}</div></div>
        <div class="ct">{item.get('tidspunkt','')}</div></div>
      <div class="cn">→ {item.get('handling','')}</div>{dl}</div>"""

def generate_html(analysis: dict, user: dict, run_time: datetime, hours_back: int,
                  alle_opgaver: list, overskredet: list, afventer: list,
                  møde_forberedelser: list, teams_msgs: list) -> tuple:
    REPORTS_DIR.mkdir(exist_ok=True)
    path = REPORTS_DIR / f"mailrapport_{run_time.strftime('%Y%m%d_%H%M')}.html"

    høj    = analysis.get("høj_prioritet", [])
    medium = analysis.get("medium_prioritet", [])
    lav    = analysis.get("lav_prioritet", [])
    deadlines = analysis.get("deadlines", [])
    åbne_opgaver = [o for o in alle_opgaver if not o.get("udført")]
    teams_vigtige = analysis.get("teams_vigtige", [])

    def cards(lst, farve, ikon):
        return "".join(_kort(m, farve, ikon) for m in lst) or '<p class="empty">Ingen i denne kategori</p>'

    # Opfølgning-sektion
    opf_html = ""
    if overskredet or afventer:
        over_rows = "".join(
            f"<tr style='background:#fff5f5'><td>🔴</td><td><strong>{o['emne']}</strong></td>"
            f"<td>{o['afsender']}</td><td style='color:#c53030'>Overskredet {OPFØLGNING_TIMER}t!</td></tr>"
            for o in overskredet
        )
        afv_rows = "".join(
            f"<tr><td>🟡</td><td>{a['emne']}</td><td>{a['afsender']}</td><td>Afventer svar</td></tr>"
            for a in afventer[:5]
        )
        opf_html = f"""<section class="sec">
          <h2>🔔 Opfølgning – Afventer svar</h2>
          <table class="tbl"><thead><tr><th></th><th>Emne</th><th>Afsender</th><th>Status</th></tr></thead>
          <tbody>{over_rows}{afv_rows}</tbody></table></section>"""

    # Møde-sektion
    møde_html = ""
    if møde_forberedelser:
        kort_list = []
        for m in møde_forberedelser:
            if m.get("all_day"):
                continue
            rel = "".join(f"<li>{r}</li>" for r in m.get("relaterede_mails", []))
            rel_html = f"<ul class='rel-list'>{rel}</ul>" if rel else "<p class='empty'>Ingen relaterede mails fundet</p>"
            kort_list.append(f"""<div class="møde-kort">
              <div class="møde-tid">{m['tidspunkt']}</div>
              <div class="møde-info">
                <strong>{m['emne']}</strong>
                <span class="møde-del">{m['deltagere']} deltagere</span>
                <div class="møde-mails">{rel_html}</div>
              </div></div>""")
        if kort_list:
            møde_html = f"""<section class="sec">
              <h2>📅 Mødeforberedelse – Næste 24 timer</h2>
              {"".join(kort_list)}</section>"""

    # Teams-sektion
    teams_html = ""
    if teams_vigtige:
        teams_rows = "".join(
            f"<div class='card' style='border-left:4px solid #6264a7'>"
            f"<div class='ch'><span>💬</span>"
            f"<div class='cm'><div class='cs'>{t.get('chat','')}</div>"
            f"<div class='ca'>fra {t.get('afsender','')}</div></div></div>"
            f"<div class='cn'>→ {t.get('besked','')}</div></div>"
            for t in teams_vigtige
        )
        teams_html = f'<section class="sec"><h2>💬 Microsoft Teams – Vigtige beskeder</h2>{teams_rows}</section>'

    # Deadline-sektion
    dl_rows = "".join(
        f"<tr><td>{d.get('emne','')}</td><td>{d.get('afsender','')}</td>"
        f"<td><strong>{d.get('deadline','')}</strong></td>"
        f"<td><span class='db'>{d.get('dage_tilbage','?') if d.get('dage_tilbage') is not None else 'Se mail'}</span></td></tr>"
        for d in deadlines
    )
    dl_html = (f"""<section class="sec"><h2>⏰ Deadlines</h2>
      <table class="tbl"><thead><tr><th>Emne</th><th>Afsender</th><th>Deadline</th><th>Tilbage</th></tr></thead>
      <tbody>{dl_rows}</tbody></table></section>""") if deadlines else ""

    # Opgave-sektion
    opg_rows = "".join(
        f"<tr><td>⬜</td><td>{o.get('opgave','')}</td>"
        f"<td>{o.get('afsender','').split('<')[0].strip()}</td>"
        f"<td>{o.get('frist','') or '–'}</td>"
        f"<td style='color:#718096;font-size:.8rem'>{o.get('dato','')}</td></tr>"
        for o in åbne_opgaver[-20:]
    )
    opg_html = (f"""<section class="sec"><h2>📋 Åbne Opgaver ({len(åbne_opgaver)})</h2>
      <table class="tbl"><thead><tr><th></th><th>Opgave</th><th>Fra</th><th>Frist</th><th>Dato</th></tr></thead>
      <tbody>{opg_rows}</tbody></table></section>""") if åbne_opgaver else ""

    stat = analysis.get("statistik", {})
    html = f"""<!DOCTYPE html><html lang="da"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Mailrapport {run_time.strftime('%d/%m/%Y %H:%M')}</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;background:#f0f4f8;color:#2d3748}}
.hdr{{background:linear-gradient(135deg,#0078d4,#005a9e);color:#fff;padding:28px 40px}}
.hdr h1{{font-size:1.6rem;font-weight:800}}.hdr .sub{{opacity:.85;margin-top:4px;font-size:.95rem}}
.hdr .usr{{margin-top:6px;font-size:.82rem;opacity:.7}}
.wrap{{max-width:960px;margin:28px auto;padding:0 20px 60px}}
.stats{{display:grid;grid-template-columns:repeat(4,1fr);gap:14px;margin-bottom:24px}}
.sc{{background:#fff;border-radius:12px;padding:18px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,.1)}}
.sc .n{{font-size:1.9rem;font-weight:800}}.sc .l{{font-size:.78rem;color:#718096;text-transform:uppercase;letter-spacing:.05em;margin-top:3px}}
.sc.r .n{{color:#e53e3e}}.sc.y .n{{color:#d69e2e}}.sc.g .n{{color:#38a169}}
.summ{{background:#fff;border-radius:12px;padding:18px 22px;margin-bottom:20px;
       box-shadow:0 1px 3px rgba(0,0,0,.1);border-left:4px solid #0078d4;font-size:.95rem;line-height:1.6}}
.sec{{background:#fff;border-radius:12px;padding:22px;margin-bottom:20px;box-shadow:0 1px 3px rgba(0,0,0,.1)}}
.sec h2{{font-size:1.05rem;margin-bottom:14px;padding-bottom:8px;border-bottom:2px solid #edf2f7}}
.card{{background:#f7fafc;border-radius:8px;padding:13px 15px;margin-bottom:9px}}
.ch{{display:flex;align-items:flex-start;gap:9px}}.cm{{flex:1}}
.cs{{font-weight:600;font-size:.93rem}}.ca{{font-size:.8rem;color:#718096;margin-top:2px}}
.ct{{font-size:.78rem;color:#a0aec0;white-space:nowrap}}
.cn{{margin-top:7px;font-size:.86rem;color:#4a5568;background:#fff;border-radius:6px;padding:5px 9px}}
.dl-badge{{display:inline-block;margin-top:7px;font-size:.78rem;background:#fff5f5;color:#c53030;
           border:1px solid #feb2b2;border-radius:20px;padding:2px 9px}}
.møde-kort{{display:flex;gap:16px;padding:12px 0;border-bottom:1px solid #edf2f7}}
.møde-kort:last-child{{border-bottom:none}}
.møde-tid{{font-size:1rem;font-weight:800;color:#0078d4;min-width:80px;padding-top:2px}}
.møde-del{{font-size:.78rem;color:#718096;margin-left:8px}}
.møde-mails{{margin-top:8px}}.rel-list{{margin-left:16px;font-size:.82rem;color:#4a5568}}
.rel-list li{{margin-bottom:3px}}
.tbl{{width:100%;border-collapse:collapse;font-size:.88rem}}
.tbl th{{text-align:left;padding:7px 12px;background:#edf2f7;color:#4a5568;font-size:.78rem;text-transform:uppercase}}
.tbl td{{padding:9px 12px;border-bottom:1px solid #edf2f7}}
.db{{background:#ebf8ff;color:#2b6cb0;border-radius:20px;padding:2px 9px;font-size:.8rem;font-weight:600}}
.empty{{color:#a0aec0;font-style:italic;font-size:.88rem;padding:6px 0}}
.foot{{text-align:center;color:#a0aec0;font-size:.78rem;margin-top:36px}}
</style></head><body>
<div class="hdr">
  <h1>📬 Outlook Mailrapport</h1>
  <div class="sub">Genereret {run_time.strftime('%d. %B %Y kl. %H:%M')} · Seneste {hours_back} timer</div>
  <div class="usr">{user.get('displayName','')} · {user.get('mail','')}</div>
</div>
<div class="wrap">
  <div class="stats">
    <div class="sc"><div class="n">{stat.get('total',len(høj)+len(medium)+len(lav))}</div><div class="l">Mails</div></div>
    <div class="sc r"><div class="n">{len(høj)}</div><div class="l">Høj prioritet</div></div>
    <div class="sc y"><div class="n">{len(medium)}</div><div class="l">Medium</div></div>
    <div class="sc g"><div class="n">{len(lav)}</div><div class="l">Lav prioritet</div></div>
  </div>
  <div class="summ"><strong>📋 Oversigt:</strong> {analysis.get('oversigt','')}</div>
  {opf_html}
  {møde_html}
  {teams_html}
  {dl_html}
  {opg_html}
  <section class="sec"><h2>🔴 Høj Prioritet</h2>{cards(høj,'#e53e3e','🔴')}</section>
  <section class="sec"><h2>🟡 Medium Prioritet</h2>{cards(medium,'#d69e2e','🟡')}</section>
  <section class="sec"><h2>🟢 Lav Prioritet</h2>{cards(lav,'#38a169','🟢')}</section>
  <div class="foot">Outlook Mailagent v3 · Claude AI · {run_time.strftime('%d/%m/%Y %H:%M')}</div>
</div></body></html>"""

    path.write_text(html, encoding="utf-8")
    return path, html


# ── Hoved ─────────────────────────────────────────────────────────────────────
def main():
    run_time = datetime.now()
    print(f"\n{'='*60}\n  📧 OUTLOOK MAILAGENT v3  –  {run_time.strftime('%d/%m/%Y %H:%M')}\n{'='*60}\n")

    missing = [k for k in ["AZURE_CLIENT_ID", "ANTHROPIC_API_KEY"] if not os.getenv(k)]
    if missing:
        print(f"❌  Manglende konfiguration: {', '.join(missing)}\n"); sys.exit(1)

    try:
        print("🔐  Checker Microsoft-login...")
        token = get_access_token()
        print("    ✓ Logget ind\n")

        user = fetch_user_info(token)
        print(f"👤  Konto: {user.get('displayName','')} ({user.get('mail','')})\n")

        # Mails
        print(f"📥  Henter mails (seneste {HOURS_BACK} timer)...")
        emails = fetch_emails(token, HOURS_BACK)
        print(f"    ✓ {len(emails)} mails\n")

        # Sendte mails (til opfølgning)
        sent = fetch_sent_emails(token, HOURS_BACK)

        # Kalender
        print("📅  Henter kalender...")
        møder = fetch_calendar_events(token)
        print(f"    ✓ {len(møder)} møder de næste 24 timer\n")

        # Teams
        print("💬  Henter Teams-beskeder...")
        teams_msgs = fetch_teams_messages(token)
        print(f"    ✓ {len(teams_msgs)} beskeder\n")

        if not emails and not teams_msgs:
            print("✅  Ingen nye mails eller Teams-beskeder\n"); return

        # Vigtige afsendere (manuelt tilføjede)
        vigtige_manuel = load_vigtige_afsendere()

        # Selvlæring – opdater scorer baseret på sendte svar
        print("🧠  Opdaterer selvlæring...")
        scorer = opdater_laering(emails, sent)
        vigtige_lært = get_laerte_vigtige(scorer)
        vigtige = list(set(vigtige_manuel + vigtige_lært))
        laering_oversigt = generer_laering_oversigt(scorer)
        print(f"    ✓ {len(vigtige_lært)} afsendere lært som vigtige\n")

        # AI-analyse
        print("🤖  Analyserer med Claude AI...")
        analysis = analyze_with_claude(emails, vigtige, teams_msgs, møder)
        print("    ✓ Analyse færdig\n")

        # Opfølgning
        overskredet, afventer = opdater_opfølgning(emails, sent, analysis)
        if overskredet:
            print(f"🔔  {len(overskredet)} mails har overskredet svarfrist!\n")

        # Mødeforberedelse
        møde_forberedelser = forbered_møder(møder, emails)

        # Opgaver
        nye_opgaver = []
        for o in analysis.get("opgaver", []):
            o["dato"] = run_time.strftime("%d/%m/%Y")
            o["udført"] = False
            nye_opgaver.append(o)
        alle_opgaver = tilføj_opgaver(nye_opgaver)
        generer_opgaver_md(alle_opgaver)
        åbne = len([o for o in alle_opgaver if not o.get("udført")])
        print(f"📋  Opgaver: {len(nye_opgaver)} nye · {åbne} åbne i alt\n")

        # Rapport
        print("📊  Genererer rapport...")
        report_path, html_content = generate_html(
            analysis, user, run_time, HOURS_BACK,
            alle_opgaver, overskredet, afventer,
            møde_forberedelser, teams_msgs
        )
        print(f"    ✓ {report_path}\n")

        # Send på mail
        if SEND_EMAIL_REPORT and user.get("mail"):
            print("📨  Sender rapport til Outlook...")
            try:
                emne = f"📬 Mailrapport {run_time.strftime('%d/%m %H:%M')} – {len(høj := analysis.get('høj_prioritet',[]))} vigtige"
                send_report_email(token, user, html_content, emne)
                print("    ✓ Sendt\n")
            except Exception as e:
                print(f"    ⚠️  Mail-fejl: {e}\n")

        print("─"*60)
        print(f"  📋 {analysis.get('oversigt','')}")
        print(f"  🔴 Høj prioritet: {len(analysis.get('høj_prioritet',[]))}")
        print(f"  🔔 Afventer svar: {len(afventer)} · Overskredet: {len(overskredet)}")
        print(f"  📅 Møder: {len(møder)}")
        print(f"  💬 Teams-beskeder: {len(teams_msgs)}")
        print(f"  📋 Nye opgaver: {len(nye_opgaver)} (åbne i alt: {åbne})")
        print(f"  🧠 Selvlæring: {laering_oversigt or 'Ikke nok data endnu'}")
        print(f"  📄 Rapport: {report_path}")
        print("─"*60 + "\n")

    except KeyboardInterrupt:
        print("\n\n⚠️  Afbrudt.")
    except requests.HTTPError as e:
        print(f"\n❌  HTTP-fejl: {e.response.status_code} – {e.response.text[:300]}"); sys.exit(1)
    except Exception as e:
        print(f"\n❌  Fejl: {e}"); raise


if __name__ == "__main__":
    main()
