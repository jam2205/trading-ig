"""
forex_epics.py
--------------
Epic configuration for the 7 Forex majors on IG Markets.

Epic format:  CS.D.{PAIR}.{TYPE}.IP
  CS        = Cash/Spot market
  D         = Daily (rolling)
  TODAY.IP  = Spreadbet daily rolling contract
  CFD.IP    = Contract for Difference

Each entry in FOREX_MAJORS contains:
  pair           raw pair code used in the Epic string
  label          human-readable pair label
  sb             Spreadbet Epic (TODAY.IP)
  cfd            CFD Epic (CFD.IP)
  news_keywords  search terms for NewsAPI.org
"""

FOREX_MAJORS = [
    {
        "pair": "EURUSD",
        "label": "EUR/USD",
        "sb": "CS.D.EURUSD.TODAY.IP",
        "cfd": "CS.D.EURUSD.CFD.IP",
        "news_keywords": "EUR USD euro dollar ECB Fed",
    },
    {
        "pair": "GBPUSD",
        "label": "GBP/USD",
        "sb": "CS.D.GBPUSD.TODAY.IP",
        "cfd": "CS.D.GBPUSD.CFD.IP",
        "news_keywords": "GBP USD pound dollar sterling BoE",
    },
    {
        "pair": "USDJPY",
        "label": "USD/JPY",
        "sb": "CS.D.USDJPY.TODAY.IP",
        "cfd": "CS.D.USDJPY.CFD.IP",
        "news_keywords": "USD JPY yen dollar BoJ",
    },
    {
        "pair": "AUDUSD",
        "label": "AUD/USD",
        "sb": "CS.D.AUDUSD.TODAY.IP",
        "cfd": "CS.D.AUDUSD.CFD.IP",
        "news_keywords": "AUD USD australian dollar RBA",
    },
    {
        "pair": "USDCAD",
        "label": "USD/CAD",
        "sb": "CS.D.USDCAD.TODAY.IP",
        "cfd": "CS.D.USDCAD.CFD.IP",
        "news_keywords": "USD CAD canadian dollar loonie BoC",
    },
    {
        "pair": "USDCHF",
        "label": "USD/CHF",
        "sb": "CS.D.USDCHF.TODAY.IP",
        "cfd": "CS.D.USDCHF.CFD.IP",
        "news_keywords": "USD CHF swiss franc SNB",
    },
    {
        "pair": "NZDUSD",
        "label": "NZD/USD",
        "sb": "CS.D.NZDUSD.TODAY.IP",
        "cfd": "CS.D.NZDUSD.CFD.IP",
        "news_keywords": "NZD USD new zealand dollar kiwi RBNZ",
    },
]

# Flat epic lists — use these to subscribe
SB_EPICS = [m["sb"] for m in FOREX_MAJORS]
CFD_EPICS = [m["cfd"] for m in FOREX_MAJORS]
ALL_EPICS = SB_EPICS + CFD_EPICS

# Lookup: epic string → metadata dict (includes "type": "Spreadbet"|"CFD")
EPIC_TO_META = {
    **{m["sb"]: {**m, "type": "Spreadbet"} for m in FOREX_MAJORS},
    **{m["cfd"]: {**m, "type": "CFD"} for m in FOREX_MAJORS},
}
