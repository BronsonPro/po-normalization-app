"""
Routing table: identifies which party a PO email belongs to, based on
sender address and a subject-line keyword. Both are checked case-insensitively.

Matching rule (see email_router.py):
  - sender must match (exact address OR domain match)
  - AND subject must contain the given keyword/phrase

If a party has multiple valid senders or subject patterns, add multiple
entries with the same party_name - the router checks all of them.
"""

PARTY_ROUTES = [
    {
        "party_name": "Reliance",
        "sender": "noreply@ril.com",
        "subject_contains": "Reliance Retail Limited",
    },
    {
        "party_name": "D-Mart",
        "sender": "narayan.dalvi@dmartindia.com",
        "subject_contains": "DMART ONLINE - ORB - IFC PO",
    },
    {
        "party_name": "D-Mart",
        "sender": "aelv-janhavi.hirlekar@dmartindia.com",
        "subject_contains": "DMART ONLINE - ORB - PO",
    },
    {
        "party_name": "D-Mart",
        "sender": "aelhofmcg@dmartindia.com",
        "subject_contains": "IFC PO",
    },
    {
        "party_name": "Manash",
        "sender": "jaya.gupta@purplle.com",
        "subject_contains": "MANASH PO",
    },
    {
        "party_name": "Manash",
        "sender": "ravi.c@purplle.com",
        "subject_contains": "MANASH PO",
    },
    {
        "party_name": "Handy Homes",
        "sender": "procurement-ops@urbancompany.com",
        "subject_contains": "Handy Homes | Purchase Orders",
    },
    {
        "party_name": "Zepto",
        "sender": "po_fulfilment@zeptonow.com",
        "subject_contains": "Purchase Order for [ORB INTERNATIONAL",
    },
    {
        "party_name": "Big Basket",
        "sender": "alerts@bigbasket.com",
        "subject_contains": "PO Details",
    },
    {
        "party_name": "Health & Glow",
        "sender": "hgbuyerblr@healthandglow.in",
        "subject_contains": "Purchase Order-",
    },
    {
        "party_name": "Blink",
        "sender": "purchaseorder@handsontrades.com",
        "subject_contains": "PO_BCPL",
    },
    {
        "party_name": "Myntra",
        "sender": "updates@myntra.com",
        "subject_contains": "New Purchase Orders",
    },
    {
        "party_name": "Scootsy",
        "sender": "no-reply.service@scootsy.com",
        "subject_contains": "NARPO",
    },
    {
        "party_name": "Nykaa",
        "sender": "noreply@nykaa.com",
        "subject_contains": "is released for ORB",
    },
    {
        "party_name": "Slikk",
        "sender": "rishit@slikk.club",
        "subject_contains": "Purchase Order",
    },
]


def identify_party(sender_address: str, subject: str):
    """
    Returns the matching party_name, or None if no route matches.
    Sender match = exact address match (case-insensitive).
    Subject match = substring match (case-insensitive).
    """
    if not sender_address or not subject:
        return None

    sender_l = sender_address.strip().lower()
    subject_l = subject.strip().lower()

    for route in PARTY_ROUTES:
        if route["sender"].lower() == sender_l:
            if route["subject_contains"].lower() in subject_l:
                return route["party_name"]

    return None