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
    "microsoftexchange329e71ec88ae4615bbc36ab6ce41109e@orbinter.com",  # auto-generated bounce notices
}

# Domains confirmed as unrelated (not a party you order from) - always IGNORED
EXCLUDED_DOMAINS = {
    "blinkit.com",
    "grofers.com",
    "reequil.com",
    "partnersbiz.com",
    "aorborc.com",
    "zohodesk.in",          # helpdesk/ticketing systems (e.g. Nykaa support tickets)
    "external.instamart.in",  # Swiggy Instamart vendor support - not a tracked party
    "delhivery.com",        # logistics/courier, not a PO source
}


# Once a party is identified by domain, this second check confirms the
# SPECIFIC EMAIL is an actual PO notification - not other correspondence
# from the same company (GRN notices, reconciliation, price updates,
# password resets, etc, which all come from the same domains and would
# otherwise pollute the review queue). A subject must contain AT LEAST ONE
# of a party's keywords to be treated as a real PO. Calibrated against a
# real batch of inbox data reviewed by hand (see PO_Fetch_Log__1_.xlsx).
PARTY_PO_KEYWORDS = {
    "Reliance": ["po. intm"],
    "D-Mart": ["dmart online - orb - ifc po", "dmart online - orb - po", "ifc po"],
    "Manash": ["manash po"],
    "Handy Homes": ["handy homes | purchase orders"],
    "Zepto": ["purchase order for [", "revised purchase order for ["],
    "Big Basket": ["po details"],
    "Health & Glow": ["purchase order-"],
    "Blink": ["po_bcpl"],
    "Myntra": ["new purchase orders"],
    "Scootsy": ["narpo"],
    "Nykaa": ["is released for orb"],
    "Slikk": ["purchase order"],
    "Zomato": ["po_bcpl"],  # Zomato sends CC'd copies of Blink's PO_BCPL threads
}


def is_po_subject(party_name: str, subject: str) -> bool:
    """True if this email's subject matches a known real-PO pattern for the party."""
    if not subject:
        return False
    subject_l = subject.strip().lower()
    keywords = PARTY_PO_KEYWORDS.get(party_name, [])
    return any(kw in subject_l for kw in keywords)


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
