"""
Routing table: identifies which party a PO email belongs to.

APPROACH: match by SENDER DOMAIN, not exact address or subject phrase.
(We initially matched on exact sender address + subject substring, but real
inbox data showed POs come from many different individual employee addresses
per company, e.g. nykaa.com POs come from sandip.patil@, purchaseorder1@,
prachi.dhanuka@, etc - and subject phrasing varies more than expected too.
Domain stays stable even as staff/subject formats change.)

Two lookups:
  PARTY_DOMAINS   - sender's domain -> party name (the real match)
  EXCLUDED_SENDERS - specific known non-PO addresses to mark IGNORED rather
                     than flagging as "needs review" noise (bank alerts,
                     internal/training mail, bounces, your own addresses)
"""

PARTY_DOMAINS = {
    "ril.com": "Reliance",
    "dmartindia.com": "D-Mart",
    "purplle.com": "Manash",
    "urbancompany.com": "Handy Homes",
    "zepto.com": "Zepto",
    "zeptonow.com": "Zepto",
    "bigbasket.com": "Big Basket",
    "healthandglow.in": "Health & Glow",
    "handsontrades.com": "Blink",
    "myntra.com": "Myntra",
    "scootsy.com": "Scootsy",
    "nykaa.com": "Nykaa",
    "slikk.club": "Slikk",
    "zomato.com": "Zomato",
}

# Specific addresses confirmed as NOT real POs - logged as IGNORED, not UNKNOWN,
# so they don't show up as something needing attention every run.
EXCLUDED_SENDERS = {
    "rr.paymentadvice@ril.com",       # payment advice, not a PO
    "bbtraining.bnm@bigbasket.com",   # internal training mail
    "sellersupport@myntra.com",       # support tickets, not POs
    "enetadvicemailing@hdfcbank.bank.in",  # bank notification
    "mailer-daemon@googlemail.com",   # bounce message
    "account@orbinter.com",           # your own address
    "orbmumbai@gmail.com",            # your own address
    "it.a.45.pravin.jaybhaye@gmail.com",  # personal gmail, not a PO source
}

# Domains confirmed as unrelated (not a party you order from) - always IGNORED
EXCLUDED_DOMAINS = {
    "blinkit.com",
    "grofers.com",
    "reequil.com",
    "partnersbiz.com",
    "aorborc.com",
}


def identify_party(sender_address: str, subject: str):
    """
    Returns:
      - party_name (str)  -> matched a known party domain
      - "IGNORE"           -> confirmed non-PO sender/domain, skip review queue
      - None                -> genuinely unrecognized, needs manual review
    """
    if not sender_address:
        return None

    sender_l = sender_address.strip().lower()

    if sender_l in EXCLUDED_SENDERS:
        return "IGNORE"

    if "@" not in sender_l:
        return None

    domain = sender_l.split("@", 1)[1]

    if domain in EXCLUDED_DOMAINS:
        return "IGNORE"

    if domain in PARTY_DOMAINS:
        return PARTY_DOMAINS[domain]

    return None
