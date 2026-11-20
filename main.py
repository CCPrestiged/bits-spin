"""
Real Humans · Mystery BITS
==========================

One file. FastAPI + SQLAlchemy + openpyxl. Nothing else, because it has to be
pasted into a web editor by hand.

Three wheels
------------
* Wheel I  Initiation  300 BITS, wager x1 .. x1,000,000
* Wheel II Tempest     3,000 BITS, wager x1 .. x10
* Wheel III Cataclysm  30,000 BITS, no wager at all

Wheel II opens after 100 finished spins on Wheel I. Wheel III opens after 1,000
finished spins on Wheel II. A spin counts when it was collected; a spin left
before the wheel stopped is a forfeit and never counts.

Rules this file implements
--------------------------
* spending BITS      = J - P * 1,000,000     (per row of the BITS Bank sheet)
* a spin costs       = the wheel's price * wager
* wager in multiples = Sovereign and above only; below that the wager is forced
                       to 1 and they simply pay the price for whatever they win
* BITS prizes scale  = a wager of 30 landing on 1,500 BITS owes 45,000
* items never scale  = logos, prefixes, fragments, boosters and tokens are
                       things, not money, so a wager leaves them alone
* fragments          = 3 of the same fragment make one Booster, automatically.
                       Fragments are kept forever; boosters are owed to the
                       player and cleared by Mark applied.
* jackpot            = every paid Wheel III spin drops its whole price into one
                       pot. The pot is paid out the moment somebody lands on the
                       jackpot, and then starts again from zero.
* reload before the spin settles = the wager is gone, the sweep keeps it
* admin              = the Lichess account in BS_ADMIN_LICHESS (CC_Prestiged)

Run:  uvicorn main:app --host 0.0.0.0 --port $PORT
"""
from __future__ import annotations

import base64
import csv
import datetime as dt
import hashlib
import hmac
import io
import json
import os
import secrets
import urllib.error
import urllib.parse
import urllib.request
import uuid

import openpyxl
from fastapi import Depends, FastAPI, File, Form, Request, UploadFile
from fastapi.responses import JSONResponse, PlainTextResponse, RedirectResponse, Response
from sqlalchemy import (BigInteger, Boolean, Column, DateTime, ForeignKey, Integer, String,
                        Text, and_, create_engine, func, inspect, select, text)
from sqlalchemy.orm import declarative_base, sessionmaker

# --------------------------------------------------------------------------- config
def env(name, default=None):
    v = os.environ.get(name)
    return v if v not in (None, "") else default


SECRET_KEY = env("BS_SECRET_KEY", "")


def normalize_db_url(url: str) -> str:
    """Hosting dashboards hand out postgres:// or postgresql://; SQLAlchemy wants a driver."""
    for prefix in ("postgres://", "postgresql://"):
        if url.startswith(prefix):
            return "postgresql+psycopg2://" + url[len(prefix):]
    return url


DB_URL = normalize_db_url(env("BS_DB_URL", "sqlite:///./bits_spin.db"))
BASE_URL = env("BS_BASE_URL", "http://127.0.0.1:8000")

ADMIN_RAW = env("BS_ADMIN_LICHESS", "CC_Prestiged")


def name_key(value) -> str:
    """One canonical form for a Lichess name: no case, no spaces, no leading @.
    Everything that compares names against BS_ADMIN_LICHESS goes through this, so
    "CC_Prestiged", "cc_prestiged" and "  @CC_Prestiged " are the same account."""
    return str(value or "").strip().lstrip("@").lower()


ADMIN_LICHESS = name_key(ADMIN_RAW)
ADMIN_KEY = env("BS_ADMIN_KEY", "")          # optional break-glass, off unless set

# Password logins are off unless you ask for them: anyone could register the name of a
# player who already has BITS. Lichess OAuth is the way in. BS_ALLOW_LOCAL=1 turns the
# password form back on, which is useful for testing on your own machine.
ALLOW_LOCAL = env("BS_ALLOW_LOCAL", "0") == "1"
BITS_HELP = env("BS_BITS_HELP",
                "https://sites.google.com/view/real-humans-codex/knowledge/bits-info-simple")
# Lichess has retired app registration (/account/oauth/app is gone) and expects public
# clients only. Their own error message for a missing client_id is "client_id required
# (choose any)", so this string is simply an identifier, not a registered application,
# and there is no secret anywhere. PKCE does the work instead.
#
#   browser  -> https://lichess.org/oauth        the consent page, with the S256 challenge
#   server   -> https://lichess.org/api/token    code -> token, once. Never /oauth.
LICHESS_AUTHORIZE = "https://lichess.org/oauth"
LICHESS_TOKEN = "https://lichess.org/api/token"
LICHESS_CLIENT_ID = env("BS_LICHESS_CLIENT_ID", "real-humans-spins")
LICHESS_REDIRECT = env("BS_LICHESS_REDIRECT", BASE_URL.rstrip("/") + "/auth/callback")

BANK_FILE = env("BS_BANK_FILE", "")
BANK_URL = env("BS_BANK_URL", "")
BANK_SHEET = env("BS_BANK_SHEET", "BITS Bank")
COL_PLAYER, COL_TOTAL, COL_SAC, COL_RANK, COL_PRESTIGE = "B", "J", "P", "K", "O"
COL_TITLE = "L"                               # column L: the hierarchical title
SAC_UNIT = 1_000_000                          # P is "BITS sacrificed, in millions"

SPIN_COST = 300                               # Wheel I, kept as a name for the price
SPIN_TTL = int(env("BS_SPIN_TTL_SECONDS", "30"))
# The wheel must actually turn before the result is revealed. Anything closed before this
# many seconds is a forfeit, which is the whole point of the timer.
REVEAL_AFTER = float(env("BS_REVEAL_SECONDS", "3.6"))
ACTIVE_SECONDS = int(env("BS_ACTIVE_SECONDS", "150"))   # a page that pinged this recently is "open"
PING_EVERY = int(env("BS_PING_SECONDS", "45"))          # how often the page says it is still there
FREE_SPIN_AWARD = 2                    # Wheel I's "2 More Spins" hands over exactly two
MAX_WAGER = int(env("BS_MAX_WAGER", "1000000"))         # hard ceiling on any wager
WAGER_MIN_RANK = env("BS_WAGER_MIN_RANK", "Sovereign")

# Wheel II opens at 100 finished Wheel I spins, Wheel III at 1000 finished Wheel II spins.
UNLOCK_SPINS = {2: 100, 3: 1000}
UNLOCK_FROM = {2: 1, 3: 2}
# The far end of Wheel III's bar. Whatever is behind it is not written down anywhere.
SECRET_TARGET = 10_000

# --------------------------------------------------------------------------- the wheels
# Every slice carries its chance in basis points, so 1/100 of a percent is the smallest
# step the code can express and the weights add up to exactly 10000 every time.

def _s(label, lines, emoji, code, kind, weight, color, ink="#2b2b2b", amount=0,
       key="", scales=False, note=""):
    return {"label": label, "lines": lines, "emoji": emoji, "code": code, "kind": kind,
            "weight": weight, "color": color, "ink": ink, "amount": amount, "key": key,
            "scales": scales, "note": note}


# --- Wheel I: the original wheel, 16 equal slices. Green and growing.
WHEEL_1 = [
    _s("50 BITS", ["50", "BITS"], "\U0001F7E2", "1f7e2", "bits", 625, "#bfe8c8", amount=50,
       scales=True),
    _s("75 BITS", ["75", "BITS"], "\U0001F331", "1f331", "bits", 625, "#a6e0b2", amount=75,
       scales=True),
    _s("100 BITS", ["100", "BITS"], "\U0001F4A7", "1f4a7", "bits", 625, "#bfe8c8", amount=100,
       scales=True),
    _s("150 BITS", ["150", "BITS"], "\U0001F539", "1f539", "bits", 625, "#a6e0b2", amount=150,
       scales=True),
    _s("200 BITS", ["200", "BITS"], "\u26A1", "26a1", "bits", 625, "#bfe8c8", amount=200,
       scales=True),
    _s("250 BITS", ["250", "BITS"], "\U0001F49C", "1f49c", "bits", 625, "#a6e0b2", amount=250,
       scales=True),
    _s("300 BITS", ["300", "BITS"], "\U0001F3AF", "1f3af", "bits", 625, "#bfe8c8", amount=300,
       scales=True),
    _s("350 BITS", ["350", "BITS"], "\U0001F36D", "1f36d", "bits", 625, "#a6e0b2", amount=350,
       scales=True),
    _s("400 BITS", ["400", "BITS"], "\U0001F4BC", "1f4bc", "bits", 625, "#bfe8c8", amount=400,
       scales=True),
    _s("450 BITS", ["450", "BITS"], "\U0001F4B0", "1f4b0", "bits", 625, "#a6e0b2", amount=450,
       scales=True),
    _s("500 BITS", ["500", "BITS"], "\U0001F3C6", "1f3c6", "bits", 625, "#bfe8c8", amount=500,
       scales=True),
    _s("1000 BITS", ["1000", "BITS"], "\U0001F451", "1f451", "bits", 625, "#f2e3a0", amount=1000,
       scales=True),
    _s("Static Logo", ["Static", "Logo"], "\U0001F5BC\uFE0F", "1f5bc", "item", 625, "#cdbce8",
       amount=1, key="static_logo", scales=True),
    _s("Moving Logo", ["Moving", "Logo"], "\U0001F39E\uFE0F", "1f39e", "item", 625, "#f3cf9b",
       amount=1, key="moving_logo", scales=True),
    _s("Custom Prefix", ["Custom", "Prefix"], "\U0001F9EC", "1f9ec", "item", 625, "#b8d8f2",
       amount=1, key="custom_prefix", scales=True),
    _s("2 More Spins", ["2 more", "spins"], "\U0001F504", "1f504", "free", 625, "#a8dfe8",
       amount=FREE_SPIN_AWARD),
]

# --- Wheel II: grey and purple, the wind between the hills. 13 slices.
WHEEL_2 = [
    _s("Bust - 100 BITS", ["BUST", "100 BITS"], "\U0001F480", "1f480", "bits", 4500, "#5f6376",
       ink="#ffe9d6", amount=100, scales=True),
    _s("600 BITS", ["600", "BITS"], "\U0001FA99", "1fa99", "bits", 1500, "#9298ad", amount=600,
       scales=True),
    _s("1,500 BITS", ["1,500", "BITS"], "\U0001F4B0", "1f4b0", "bits", 1500, "#a9a3c9",
       amount=1500, scales=True),
    _s("3,000 BITS", ["3,000", "BITS"], "\U0001F4B5", "1f4b5", "bits", 1000, "#bcb4e0",
       amount=3000, scales=True),
    _s("6,000 BITS", ["6,000", "BITS"], "\U0001F48E", "1f48e", "bits", 600, "#8b83b8",
       ink="#ffe9d6", amount=6000, scales=True),
    _s("15,000 BITS", ["15,000", "BITS"], "\U0001F4A0", "1f4a0", "bits", 400, "#cfc7f0",
       amount=15000, scales=True),
    _s("60,000 BITS", ["60,000", "BITS"], "\U0001F3C6", "1f3c6", "bits", 45, "#6f659a",
       ink="#ffe9d6", amount=60000, scales=True),
    _s("150,000 BITS", ["150,000", "BITS"], "\U0001F451", "1f451", "bits", 5, "#e2dcf7",
       amount=150000, scales=True),
    _s("5 More Spins", ["5 more", "spins"], "\U0001F504", "1f504", "free", 50, "#b6c1d6",
       amount=5),
    _s("x2 BITS fragment", ["x2", "fragment"], "\u26A1", "26a1", "item", 220, "#9589c2",
       amount=1, key="frag_x2"),
    _s("x3 BITS fragment", ["x3", "fragment"], "\U0001F525", "1f525", "item", 100, "#a79ad6",
       amount=1, key="frag_x3"),
    _s("x5 BITS fragment", ["x5", "fragment"], "\U0001F680", "1f680", "item", 60, "#b9aee6",
       amount=1, key="frag_x5"),
    _s("x10 BITS fragment", ["x10", "fragment"], "\U0001F31F", "1f31f", "item", 20, "#cdc4f2",
       amount=1, key="frag_x10"),
]

# --- Wheel III: fire, and no wagering. 18 slices.
WHEEL_3 = [
    _s("Bust - 0 BITS", ["BUST", "0 BITS"], "\U0001F480", "1f480", "bits", 6000, "#33100f",
       ink="#f0cfc4", amount=0),
    _s("1,500 BITS", ["1,500", "BITS"], "\U0001F4B0", "1f4b0", "bits", 1000, "#7d2a17",
       ink="#ffe9d6", amount=1500),
    _s("6,000 BITS", ["6,000", "BITS"], "\U0001F48E", "1f48e", "bits", 800, "#9c3719",
       ink="#ffe9d6", amount=6000),
    _s("15,000 BITS", ["15,000", "BITS"], "\U0001F4A0", "1f4a0", "bits", 700, "#bd4a1d",
       ink="#fff0dc", amount=15000),
    _s("30,000 BITS", ["30,000", "BITS"], "\U0001F3C6", "1f3c6", "bits", 500, "#d96820",
       ink="#2a1200", amount=30000),
    _s("90,000 BITS", ["90,000", "BITS"], "\U0001F451", "1f451", "bits", 400, "#ec8f2c",
       ink="#2a1200", amount=90000),
    _s("300,000 BITS", ["300,000", "BITS"], "\U0001F409", "1f409", "bits", 200, "#f2b03c",
       ink="#2a1200", amount=300000),
    _s("3,000,000 BITS", ["3,000,000", "BITS"], "\U0001F30C", "1f30c", "bits", 20, "#ffe6a6",
       ink="#2a1200", amount=3000000),
    _s("Progressive Legacy Jackpot", ["JACKPOT", "the whole pot"], "\U0001F3B0", "1f3b0",
       "jackpot", 5, "#5d1130", ink="#ffd98a", amount=0),
    _s("1 Legacy Token", ["1", "Token"], "\U0001FA99", "1fa99", "item", 110, "#8a4a1b",
       ink="#ffe9d6", amount=1, key="legacy_token"),
    _s("3 Legacy Tokens", ["3", "Tokens"], "\U0001F45B", "1f45b", "item", 50, "#a25c21",
       ink="#fff0dc", amount=3, key="legacy_token"),
    _s("5 Legacy Tokens", ["5", "Tokens"], "\U0001FA8E", "1fa8e", "item", 15, "#b97026",
       ink="#2a1200", amount=5, key="legacy_token"),
    _s("x3 BITS Booster", ["x3", "BITS"], "\u26A1", "26a1", "item", 60, "#6b1b2c",
       ink="#ffe1d2", amount=1, key="bits_x3"),
    _s("x5 BITS Booster", ["x5", "BITS"], "\U0001F525", "1f525", "item", 30, "#822338",
       ink="#ffe1d2", amount=1, key="bits_x5"),
    _s("x10 BITS Booster", ["x10", "BITS"], "\U0001F680", "1f680", "item", 20, "#9a2a44",
       ink="#ffe1d2", amount=1, key="bits_x10"),
    _s("x20 BITS Booster", ["x20", "BITS"], "\u2604\uFE0F", "2604", "item", 10, "#b23250",
       ink="#ffe9d6", amount=1, key="bits_x20"),
    _s("x3 Legacy Booster", ["x3", "LEGACY"], "\U0001F6E1\uFE0F", "1f6e1", "item", 60, "#3a2a4e",
       ink="#e3d8f4", amount=1, key="legacy_x3"),
    _s("x5 Legacy Booster", ["x5", "LEGACY"], "\U0001F52E", "1f52e", "item", 20, "#4b3764",
       ink="#eee6fa", amount=1, key="legacy_x5"),
]

WHEELS = {
    1: {"id": 1, "numeral": "I", "name": "Initiation", "cost": SPIN_COST, "max_wager": MAX_WAGER,
        "wager": True, "slices": WHEEL_1,
        "tag": "Green, calm, and cheap. Every spin is 300 BITS times your wager.",
        "mood": "It grows. Nothing here is in a hurry."},
    2: {"id": 2, "numeral": "II", "name": "Tempest", "cost": 3000, "max_wager": min(10, MAX_WAGER),
        "wager": True, "slices": WHEEL_2,
        "tag": "Wind and grey rain. 3,000 BITS a spin, wagers up to x10.",
        "mood": "The wind takes as much as it gives, and it takes more often."},
    3: {"id": 3, "numeral": "III", "name": "Cataclysm", "cost": 30_000, "max_wager": 1,
        "wager": False, "slices": WHEEL_3,
        "tag": "Ash, ember and no wagering. 30,000 BITS a spin, and the whole price "
               "goes into the jackpot.",
        "mood": "Nothing about this wheel is steady."},
}

WHEEL = WHEEL_1                    # the old name for Wheel I, still used in a few places

# Every item a slice can hand over. The keys are also the CSV column names.
ITEMS = {
    "static_logo": {"name": "Static Logo", "column": "static_logos",
                    "emoji": "\U0001F5BC\uFE0F", "code": "1f5bc"},
    "moving_logo": {"name": "Moving Logo", "column": "moving_logos",
                    "emoji": "\U0001F39E\uFE0F", "code": "1f39e"},
    "custom_prefix": {"name": "Custom Prefix", "column": "custom_prefixes",
                      "emoji": "\U0001F9EC", "code": "1f9ec"},
    "bits_x2": {"name": "x2 BITS Booster", "column": "bits_x2_booster", "emoji": "\u26A1",
                "code": "26a1"},
    "bits_x3": {"name": "x3 BITS Booster", "column": "bits_x3_booster", "emoji": "\U0001F525",
                "code": "1f525"},
    "bits_x5": {"name": "x5 BITS Booster", "column": "bits_x5_booster", "emoji": "\U0001F680",
                "code": "1f680"},
    "bits_x10": {"name": "x10 BITS Booster", "column": "bits_x10_booster",
                 "emoji": "\U0001F31F", "code": "1f31f"},
    "bits_x20": {"name": "x20 BITS Booster", "column": "bits_x20_booster",
                 "emoji": "\u2604\uFE0F", "code": "2604"},
    "legacy_x3": {"name": "x3 Legacy Booster", "column": "legacy_x3_booster",
                  "emoji": "\U0001F6E1\uFE0F", "code": "1f6e1"},
    "legacy_x5": {"name": "x5 Legacy Booster", "column": "legacy_x5_booster",
                  "emoji": "\U0001F52E", "code": "1f52e"},
    "legacy_token": {"name": "Legacy Token", "column": "legacy_tokens", "emoji": "\U0001FA99",
                     "code": "1fa99"},
}
# Columns H .. R of the export, in order.
PRIZE_KEYS = ["static_logo", "moving_logo", "custom_prefix",
              "bits_x2", "bits_x3", "bits_x5", "bits_x10", "bits_x20",
              "legacy_x3", "legacy_x5", "legacy_token"]
# H, I, J stay the three things the sheet has always had; K to R are the boosters.
SHEET_PRIZE_COLUMNS = ["static_logo", "moving_logo", "custom_prefix"]

# A fragment turns into the booster of the same name, three at a time.
FRAGMENTS = {
    "frag_x2": "bits_x2",
    "frag_x3": "bits_x3",
    "frag_x5": "bits_x5",
    "frag_x10": "bits_x10",
}
FRAGMENT_ORDER = ["frag_x2", "frag_x3", "frag_x5", "frag_x10"]
FRAGMENTS_PER_BOOSTER = 3

# --------------------------------------------------------------------------- the ladders
# Taken from the Legacy Bank sheet of the BITS Bank workbook, top to bottom.
LEGACY_LADDER = ["Organiser", "Herald", "Sentinel", "Enforcer", "Sovereign", "Configurator",
                 "Administrator", "Bronze", "Silver", "Gold", "Platinum", "Osmium",
                 "Titanium", "Obsidian", "Diamond", "Founder"]
TITLE_LADDER = ["Aspirant", "Outer Member", "Inner Member", "Core Member", "Steward",
                "Manager", "Leader", "Elder", "Patriarch", "Ancestor",
                "Immortal Disciple", "Heavenly Guardian", "Celestial Strategist",
                "Eternal Sage", "Void Wanderer", "Star Forger", "Apex Deity",
                "Primordial Origin", "Netherfounder"]


def norm(text) -> str:
    return "".join(ch for ch in str(text or "").lower() if "a" <= ch <= "z")


def _ladder_index(name, ladder) -> int:
    """The sheet is hand typed: match exactly, else the longest ladder name inside."""
    key = norm(name)
    if not key:
        return -1
    exact = {norm(step): i for i, step in enumerate(ladder)}
    if key in exact:
        return exact[key]
    hits = [(len(norm(s)), i) for i, s in enumerate(ladder) if norm(s) and norm(s) in key]
    return max(hits)[1] if hits else -1


def rank_index(name) -> int:
    return _ladder_index(name, LEGACY_LADDER)


def title_index(name) -> int:
    return _ladder_index(name, TITLE_LADDER)


def rank_of(name) -> str:
    i = rank_index(name)
    return LEGACY_LADDER[i] if i >= 0 else ""


def rank_ordinal(rank_name, title_name="") -> int:
    """Where this player sits on the two ladders, as one sortable number.

    Rank first, then the hierarchical title, so two Bronzes are ordered by title
    rather than by the alphabet. Somebody the sheet does not recognise sorts at the
    very bottom, which is where a blank rank belongs.
    """
    r, t = rank_index(rank_name), title_index(title_name)
    return (r + 1) * 100 + (t + 1)


def can_wager_multiple(rank_name) -> bool:
    return rank_index(rank_name) >= rank_index(WAGER_MIN_RANK)


def utcnow():
    return dt.datetime.now(dt.timezone.utc)

# --------------------------------------------------------------------------- database
Base = declarative_base()


class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True)
    lichess_id = Column(String(64), unique=True, index=True)
    display = Column(String(64))
    password_hash = Column(String(255))
    approved = Column(Boolean, default=False)   # vestigial: nobody is gated on this now
    is_admin = Column(Boolean, default=False)
    created_at = Column(DateTime(timezone=True), default=utcnow)
    bank_bits = Column(BigInteger, default=0)        # J - P*1e6 at the last sync
    bank_rank = Column(String(64), default="")
    bank_total = Column(BigInteger, default=0)
    prestige = Column(Integer, default=0)
    bank_synced_at = Column(DateTime(timezone=True))
    applied_through = Column(Integer, default=0)
    last_seen = Column(DateTime(timezone=True))     # last time this player had the page open
    test_unlocked = Column(Boolean, default=False)  # admin switch: every wheel open at once


class Ledger(Base):
    __tablename__ = "ledger"
    id = Column(Integer, primary_key=True)
    user_id = Column(Integer, ForeignKey("users.id"), index=True)
    spin_id = Column(String(36), index=True)
    delta = Column(BigInteger)
    reason = Column(String(24))
    created_at = Column(DateTime(timezone=True), default=utcnow)


class Spin(Base):
    __tablename__ = "spins"
    id = Column(String(36), primary_key=True)
    user_id = Column(Integer, ForeignKey("users.id"), index=True)
    wheel = Column(Integer, default=1)               # 1, 2 or 3
    wager = Column(Integer)
    capped = Column(Boolean, default=False)          # wager forced down to 1
    cost = Column(BigInteger)
    status = Column(String(12), default="pending")   # pending | revealed | settled | forfeited
    commit = Column(String(64))
    seed = Column(String(64))
    index_landed = Column(Integer)
    label = Column(String(48))
    kind = Column(String(8))                         # bits | item | free | jackpot
    payout = Column(BigInteger, default=0)           # BITS paid
    prize_qty = Column(BigInteger, default=0)        # how many logos / spins, etc.
    note = Column(String(48), default="")            # "2/3" on a fragment, "+1 booster" etc.
    created_at = Column(DateTime(timezone=True), default=utcnow)
    expires_at = Column(DateTime(timezone=True))
    settled_at = Column(DateTime(timezone=True))


class FreeSpin(Base):
    """Free spins are a credit, not a prize to hand out: `remaining` spins at `wager`,
    queued in the order they were won, and they stay put until they are used.

    They belong to the wheel that paid them out, so a Wheel II free spin can only be
    spent on Wheel II at the wager it was won at."""
    __tablename__ = "free_spins"
    id = Column(Integer, primary_key=True)
    user_id = Column(Integer, ForeignKey("users.id"), index=True)
    wheel = Column(Integer, default=1)
    wager = Column(Integer)
    remaining = Column(Integer, default=0)
    created_at = Column(DateTime(timezone=True), default=utcnow)


class Fragment(Base):
    """Fragment counts never reset. Three of a kind quietly become a Booster."""
    __tablename__ = "fragments"
    id = Column(Integer, primary_key=True)
    user_id = Column(Integer, ForeignKey("users.id"), index=True)
    key = Column(String(24))
    count = Column(Integer, default=0)
    created_at = Column(DateTime(timezone=True), default=utcnow)


class Award(Base):
    """Every item a player has won and has not been handed yet. This is the only
    source of the prize obligations list, so one table answers for all three wheels."""
    __tablename__ = "awards"
    id = Column(Integer, primary_key=True)
    user_id = Column(Integer, ForeignKey("users.id"), index=True)
    spin_id = Column(String(36), index=True)
    key = Column(String(24))
    qty = Column(BigInteger, default=0)
    note = Column(String(48), default="")
    created_at = Column(DateTime(timezone=True), default=utcnow)


class Meta(Base):
    __tablename__ = "meta"
    k = Column(String(48), primary_key=True)
    v = Column(Text)


engine = create_engine(DB_URL, future=True,
                       **({"connect_args": {"check_same_thread": False}}
                          if DB_URL.startswith("sqlite") else {}))
SessionLocal = sessionmaker(bind=engine, expire_on_commit=False, future=True)
Base.metadata.create_all(engine)


def touch(db, user, force=False):
    """Remember that this player has the page open.

    Written at most once every 20 seconds per player, so a busy page does not turn into a
    busy database.
    """
    if not user:
        return
    now = utcnow()
    last = _aware(user.last_seen) if user.last_seen else None
    if not force and last and (now - last).total_seconds() < 20:
        return
    user.last_seen = now
    db.commit()


def is_active(user, now=None):
    if not user or not user.last_seen:
        return False
    now = now or utcnow()
    return (now - _aware(user.last_seen)).total_seconds() <= ACTIVE_SECONDS


def seen_iso(user):
    return _aware(user.last_seen).isoformat() if user and user.last_seen else ""


def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def delta_for(db, user_id, after_id=0) -> int:
    q = select(func.coalesce(func.sum(Ledger.delta), 0)).where(
        Ledger.user_id == user_id, Ledger.id > after_id)
    return int(db.execute(q).scalar() or 0)


def available_bits(db, user) -> int:
    return int(user.bank_bits or 0) + delta_for(db, user.id, user.applied_through or 0)


def meta_get(db, key, default=""):
    row = db.get(Meta, key)
    return row.v if row else default


def meta_set(db, key, value):
    db.merge(Meta(k=key, v=value))


# --------------------------------------------------------------------------- free spins
def free_grants(db, user_id, wheel=None):
    """Oldest grant first. `wheel` limits it to the wheel that is about to be spun."""
    q = select(FreeSpin).where(FreeSpin.user_id == user_id, FreeSpin.remaining > 0)
    if wheel is not None:
        q = q.where(FreeSpin.wheel == int(wheel))
    return db.execute(q.order_by(FreeSpin.created_at, FreeSpin.id)).scalars().all()


def free_state(db, user, wheel=None):
    grants = free_grants(db, user.id, wheel)
    if not grants:
        return None
    head = grants[0]
    return {"count": head.remaining, "wager": head.wager, "wheel": int(head.wheel or 1),
            "total": sum(g.remaining for g in grants),
            "queue": [{"wager": g.wager, "count": g.remaining} for g in grants]}


def free_state_all(db, user):
    """One entry per wheel, so a page can show the pile that belongs to the wheel it is on
    without a round trip every time somebody switches tabs."""
    out = {}
    for w in (1, 2, 3):
        state = free_state(db, user, w)
        if state:
            out[str(w)] = state
    return out


def grant_free_spins(db, user_id, wheel, wager, count=FREE_SPIN_AWARD):
    row = db.execute(select(FreeSpin)
                     .where(FreeSpin.user_id == user_id, FreeSpin.wheel == int(wheel),
                            FreeSpin.wager == wager)
                     .order_by(FreeSpin.created_at)).scalars().first()
    if row:
        row.remaining += count
    else:
        db.add(FreeSpin(user_id=user_id, wheel=int(wheel), wager=wager, remaining=count))


# --------------------------------------------------------------------------- the jackpot
def jackpot_value(db) -> int:
    try:
        return int(meta_get(db, "jackpot", "0") or 0)
    except ValueError:
        return 0


def jackpot_add(db, amount):
    if amount:
        meta_set(db, "jackpot", str(jackpot_value(db) + int(amount)))


def jackpot_take(db) -> int:
    value = jackpot_value(db)
    meta_set(db, "jackpot", "0")
    return value


# --------------------------------------------------------------------------- fragments
def fragments_of(db, user_id) -> dict:
    rows = db.execute(select(Fragment).where(Fragment.user_id == user_id)).scalars().all()
    return {r.key: int(r.count or 0) for r in rows if (r.count or 0) > 0}


def give_fragment(db, user_id, key, count=1, spin_id=None):
    """Add fragments and hand over a Booster for every three of them.

    Returns (boosters_made, fragments_left_over).
    """
    row = db.execute(select(Fragment).where(Fragment.user_id == user_id,
                                            Fragment.key == key)).scalars().first()
    if not row:
        row = Fragment(user_id=user_id, key=key, count=0)
        db.add(row)
        db.flush()
    row.count = int(row.count or 0) + int(count)
    made = 0
    while row.count >= FRAGMENTS_PER_BOOSTER:
        row.count -= FRAGMENTS_PER_BOOSTER
        made += 1
    if made:
        give_award(db, user_id, FRAGMENTS[key], made, spin_id=spin_id, note="3 fragments")
    return made, row.count


def give_award(db, user_id, key, qty, spin_id=None, note=""):
    if not qty:
        return None
    award = Award(user_id=user_id, spin_id=spin_id, key=key, qty=int(qty), note=note)
    db.add(award)
    return award


# --------------------------------------------------------------------------- the wheels
def slices(wheel) -> list:
    return WHEELS[int(wheel)]["slices"]


def wheel_cost(wheel) -> int:
    return int(WHEELS[int(wheel)]["cost"])


def wheel_max_wager(wheel) -> int:
    return max(1, int(WHEELS[int(wheel)]["max_wager"]))


def bp_total(wheel) -> int:
    return sum(int(s["weight"]) for s in slices(wheel))


def commit_of(seed: str) -> str:
    return hashlib.sha256(seed.encode("ascii")).hexdigest()


def pick_index(wheel, seed: str) -> int:
    """Where this seed stops.

    The weights are basis points, so the total is always 10000 and a seed is sliced
    against that. A 256-bit seed mapped onto ten thousand outcomes leaves no meaningful
    bias; the outcome depends only on the seed, so anybody can check it afterwards.
    """
    total = bp_total(wheel)
    n = int(seed, 16) % total
    run = 0
    for i, slice_ in enumerate(slices(wheel)):
        run += int(slice_["weight"])
        if n < run:
            return i
    return len(slices(wheel)) - 1


def average_return(wheel) -> dict:
    """What a spin hands back in BITS on average, over a great many spins.

    Items and free spins are left out, because they are not BITS.
    """
    total = sum(int(s["weight"]) for s in slices(wheel)) or 1
    bits = 0.0
    for s in slices(wheel):
        if s["kind"] == "bits":
            bits += int(s["amount"]) * int(s["weight"])
    per = bits / total
    return {"per_spin": per, "cost": wheel_cost(wheel),
            "pct": (per / wheel_cost(wheel) * 100.0) if wheel_cost(wheel) else 0.0}


# --------------------------------------------------------------------------- unlock
def wheel_counts(db, user_id) -> dict:
    """Finished spins per wheel. A forfeit does not count, and neither does a spin that
    is still turning."""
    rows = db.execute(select(Spin.wheel, func.count(Spin.id))
                      .where(Spin.user_id == user_id, Spin.status == "settled")
                      .group_by(Spin.wheel)).all()
    out = {1: 0, 2: 0, 3: 0}
    for wheel, count in rows:
        out[int(wheel or 1)] = out.get(int(wheel or 1), 0) + int(count or 0)
    return out


def unlocked_wheels(db, user, counts=None) -> set:
    counts = counts if counts is not None else wheel_counts(db, user.id)
    have = {1}
    if counts.get(1, 0) >= UNLOCK_SPINS[2]:
        have.add(2)
    if counts.get(2, 0) >= UNLOCK_SPINS[3]:
        have.add(3)
    if user is not None and user.test_unlocked:
        return {1, 2, 3}
    return have


def secret_progress(db, user_id) -> int:
    """BITS won on Wheel III towards the thing at the end of the bar."""
    return int(db.execute(select(func.coalesce(func.sum(Spin.payout), 0))
                          .where(Spin.user_id == user_id, Spin.wheel == 3,
                                 Spin.status == "settled")).scalar() or 0)


def progress_for(db, user, wheel, counts=None):
    """The glowing bar under the wheel: what it is counting, and how far along it is."""
    counts = counts if counts is not None else wheel_counts(db, user.id)
    if int(wheel) == 1:
        return {"have": counts.get(1, 0), "need": UNLOCK_SPINS[2], "unit": "spins",
                "target": "Wheel II", "secret": False}
    if int(wheel) == 2:
        return {"have": counts.get(2, 0), "need": UNLOCK_SPINS[3], "unit": "spins",
                "target": "Wheel III", "secret": False}
    return {"have": secret_progress(db, user.id), "need": SECRET_TARGET,
            "unit": "BITS won here", "target": "???", "secret": True}


# --------------------------------------------------------------------------- the spin
class Refused(Exception):
    def __init__(self, message):
        super().__init__(message)
        self.message = message


def _aware(value):
    if value is None:
        return None
    return value if value.tzinfo else value.replace(tzinfo=dt.timezone.utc)


def sweep(db, now=None) -> int:
    now = now or utcnow()
    n = 0
    live = select(Spin).where(Spin.status.in_(("pending", "revealed")))
    for s in db.execute(live).scalars().all():
        if _aware(s.expires_at) and _aware(s.expires_at) <= now:
            s.status = "forfeited"
            s.settled_at = now
            n += 1
    if n:
        db.commit()
    return n


def start_spin(db, user, wager, free=False, wheel=1):
    sweep(db)
    wheel = int(wheel or 1)
    if wheel not in WHEELS:
        raise Refused("There is no such wheel.")
    spec = WHEELS[wheel]
    if wheel not in unlocked_wheels(db, user):
        need = UNLOCK_FROM.get(wheel)
        if need:
            have = wheel_counts(db, user.id).get(need, 0)
            raise Refused("Wheel %s is locked. %s of %s spins on Wheel %s so far."
                          % (spec["numeral"], f"{have:,}", f"{UNLOCK_SPINS[wheel]:,}",
                             WHEELS[need]["numeral"]))
        raise Refused("That wheel is locked.")

    capped = False
    if free:
        grants = free_grants(db, user.id, wheel)
        if not grants:
            raise Refused("You have no free spins on this wheel.")
        grant = grants[0]
        wager, cost = grant.wager, 0     # the wager it was won at, and it costs nothing
        grant.remaining -= 1
        if grant.remaining <= 0:
            db.delete(grant)
        balance = available_bits(db, user)
    else:
        try:
            wager = int(wager)
        except (TypeError, ValueError):
            raise Refused("Pick a whole number wager.")
        if not spec["wager"]:
            wager = 1                     # Wheel III has no wager at all
        else:
            multiple = can_wager_multiple(user.bank_rank)
            if not multiple:
                wager, capped = 1, True   # below Sovereign: one price, no multiple
            max_wager = wheel_max_wager(wheel)
            if wager < 1 or wager > max_wager:
                raise Refused("The wager on Wheel %s runs from 1 to %s."
                              % (spec["numeral"], f"{max_wager:,}"))

        cost = wheel_cost(wheel) * max(1, wager)
        balance = available_bits(db, user)
        if cost > balance:
            if balance <= 0:
                raise Refused("You have no BITS yet, so there is nothing to spin with. "
                              "How BITS work: " + BITS_HELP)
            raise Refused("That costs %s BITS; you have %s." % (f"{cost:,}", f"{balance:,}"))

    seed = secrets.token_hex(32)
    spin = Spin(id=str(uuid.uuid4()), user_id=user.id, wheel=wheel, wager=max(1, wager),
                capped=capped, cost=cost, status="pending", commit=commit_of(seed), seed=seed,
                created_at=utcnow(), expires_at=utcnow() + dt.timedelta(seconds=SPIN_TTL))
    db.add(spin)
    if cost:
        db.add(Ledger(user_id=user.id, spin_id=spin.id, delta=-cost, reason="spin"))
        if wheel == 3:
            # the whole price of a Cataclysm spin drops into the shared pot
            jackpot_add(db, cost)
    db.commit()
    return spin, cost, balance - cost


def reveal_spin(db, user, spin_id, now=None):
    """Work out where the wheel stops, without paying anything out yet.

    Only allowed once the wheel has had time to turn: a page that is closed before this
    never sees a result, and the spin is swept away as a forfeit."""
    now = now or utcnow()
    sweep(db, now)
    spin = db.get(Spin, spin_id)
    if not spin or spin.user_id != user.id:
        raise Refused("That spin does not exist.")
    if spin.status == "settled":
        return spin
    if spin.status in ("forfeited",):
        raise Refused("That spin expired before it was finished, so the wager is gone.")
    if spin.status != "revealed":
        ready = _aware(spin.created_at) + dt.timedelta(seconds=REVEAL_AFTER)
        if now < ready:
            raise Refused("The wheel is still turning.")
        wheel = int(spin.wheel or 1)
        index = pick_index(wheel, spin.seed)
        slice_ = slices(wheel)[index]
        spin.index_landed, spin.label, spin.kind = index, slice_["label"], slice_["kind"]
        if slice_["kind"] == "bits":
            spin.payout = int(slice_["amount"]) * int(spin.wager)
        elif slice_["kind"] in ("free", "jackpot"):
            spin.payout, spin.prize_qty = 0, 0      # settled at claim
        else:
            # money scales with the wager; things never do
            spin.payout = 0
            spin.prize_qty = (int(slice_["amount"]) * int(spin.wager) if slice_["scales"]
                              else int(slice_["amount"]))
        spin.status = "revealed"
        db.commit()
    return spin


def claim_spin(db, user, spin_id, now=None):
    """Hand over the prize. Only a spin that was revealed, and is still alive, pays."""
    now = now or utcnow()
    sweep(db, now)
    spin = db.get(Spin, spin_id)
    if not spin or spin.user_id != user.id:
        raise Refused("That spin does not exist.")
    if spin.status == "settled":
        return spin, False
    if spin.status == "forfeited":
        raise Refused("That spin expired before it was collected, so nothing was paid.")
    if spin.status != "revealed":
        raise Refused("The wheel has not stopped yet.")
    wheel = int(spin.wheel or 1)

    if spin.kind == "jackpot":
        spin.payout = jackpot_take(db)      # the pot as it stands at this moment
        if spin.payout:
            db.add(Ledger(user_id=user.id, spin_id=spin.id, delta=spin.payout,
                          reason="jackpot"))
    elif spin.payout:
        db.add(Ledger(user_id=user.id, spin_id=spin.id, delta=spin.payout, reason="payout"))

    if spin.kind == "free":
        count = _free_award(wheel, spin)
        grant_free_spins(db, user.id, wheel, spin.wager, count)
        spin.prize_qty = count                   # what the claim actually handed over
    elif spin.kind in ("item", "prize"):
        slice_ = _slice_landed(spin)
        key = (slice_ or {}).get("key") or _key_for_label(spin.label)
        qty = int(spin.prize_qty or 0)
        if key in FRAGMENTS:
            made, left = give_fragment(db, user.id, key, qty, spin_id=spin.id)
            spin.note = ("+%d booster" % made) if made else ("%d/%d" % (
                left, FRAGMENTS_PER_BOOSTER))
        elif key:
            give_award(db, user.id, key, qty, spin_id=spin.id)

    spin.status = "settled"
    spin.settled_at = now
    db.commit()
    return spin, True


def _slice_landed(spin):
    try:
        wheel = int(spin.wheel or 1)
        if spin.index_landed is not None and 0 <= spin.index_landed < len(slices(wheel)):
            return slices(wheel)[spin.index_landed]
    except (KeyError, TypeError, ValueError):
        pass
    return None


def _free_award(wheel, spin):
    slice_ = _slice_landed(spin)
    if slice_ and slice_["kind"] == "free":
        return int(slice_["amount"])
    return FREE_SPIN_AWARD


def _key_for_label(label):
    """Old rows only stored the label. This is how they find their item again."""
    for key, item in ITEMS.items():
        if item["name"] == label:
            return key
    for slice_ in WHEEL_1:
        if slice_["label"] == label and slice_["key"]:
            return slice_["key"]
    return ""

# --------------------------------------------------------------------------- upgrades
def ensure_columns():
    """Add the columns this version needs to a database that was created before them.

    A site like this is deployed once and then kept, so the tables are patched in place
    rather than dropped and rebuilt. Everything here is additive and runs once.
    """
    stamp = "TIMESTAMP WITH TIME ZONE" if engine.dialect.name != "sqlite" else "TIMESTAMP"
    wanted = [
        ("users", "last_seen", stamp),
        ("users", "test_unlocked", "BOOLEAN DEFAULT 0"),
        ("spins", "wheel", "INTEGER DEFAULT 1"),
        ("spins", "note", "VARCHAR(48) DEFAULT ''"),
        ("free_spins", "wheel", "INTEGER DEFAULT 1"),
    ]
    try:
        have = {}
        for table, column, ddl in wanted:
            if table not in have:
                have[table] = {c["name"] for c in inspect(engine).get_columns(table)}
            if column in have[table]:
                continue
            with engine.begin() as conn:
                conn.execute(text("ALTER TABLE %s ADD COLUMN %s %s" % (table, column, ddl)))
            have[table].add(column)
    except Exception:                                            # noqa: BLE001
        return


def backfill_awards(db):
    """Items won before the awards table existed are still owed. Give each old prize
    row an award of its own, once."""
    try:
        rows = db.execute(select(Spin).where(Spin.prize_qty > 0,
                                            Spin.status == "settled",
                                            Spin.kind == "prize")).scalars().all()
    except Exception:                                            # noqa: BLE001
        return 0
    if not rows:
        return 0
    known = {sid for (sid,) in db.execute(select(Award.spin_id)).all() if sid}
    made = 0
    for spin in rows:
        if spin.id in known:
            continue
        key = _key_for_label(spin.label)
        if not key:
            continue
        db.add(Award(user_id=spin.user_id, spin_id=spin.id, key=key,
                     qty=int(spin.prize_qty or 0), note="before awards",
                     created_at=spin.settled_at or spin.created_at))
        made += 1
    if made:
        db.commit()
    return made


# --------------------------------------------------------------------------- bank
class BankError(Exception):
    pass


def col_index(letter: str) -> int:
    n = 0
    for ch in letter.strip().upper():
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def _download(url: str) -> bytes:
    if any(h in url for h in ("1drv.ms", "onedrive.live.com", "sharepoint.com")):
        url += ("&" if "?" in url else "?") + "download=1"
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 (Real Humans BITS)"})
    with urllib.request.urlopen(req, timeout=45) as r:
        blob = r.read()
    if not blob[:2] == b"PK":
        raise BankError("The link returned a web page, not the workbook. Upload the file instead.")
    return blob


def fetch_bank() -> tuple[bytes, str]:
    if BANK_FILE and os.path.exists(BANK_FILE):
        with open(BANK_FILE, "rb") as fh:
            return fh.read(), "file:" + BANK_FILE
    if BANK_URL:
        return _download(BANK_URL), "url:" + BANK_URL
    raise BankError("No bank configured. Set BS_BANK_FILE or BS_BANK_URL, or upload the file.")


def parse_bank(blob: bytes) -> list[dict]:
    wb = openpyxl.load_workbook(io.BytesIO(blob), data_only=True, read_only=True)
    if BANK_SHEET not in wb.sheetnames:
        raise BankError("No sheet called %r in that workbook." % BANK_SHEET)
    ws = wb[BANK_SHEET]
    i_player, i_total = col_index(COL_PLAYER), col_index(COL_TOTAL)
    i_sac, i_rank, i_p = col_index(COL_SAC), col_index(COL_RANK), col_index(COL_PRESTIGE)
    i_title = col_index(COL_TITLE)
    rows = []
    for raw in ws.iter_rows(values_only=True):
        name = raw[i_player] if len(raw) > i_player else None
        if not isinstance(name, str) or not name.strip().startswith("@"):
            continue
        display = name.strip().lstrip("@")
        if not display or display.lower() == "username":
            continue
        total = raw[i_total] if len(raw) > i_total else 0
        sac = raw[i_sac] if len(raw) > i_sac else 0
        try:
            total = int(round(float(total or 0)))
        except (TypeError, ValueError):
            total = 0
        try:
            sac = int(round(float(sac or 0)))
        except (TypeError, ValueError):
            sac = 0
        prestige = int(raw[i_p] or 0) if len(raw) > i_p and isinstance(raw[i_p], (int, float)) else 0
        if not sac and prestige > 0:
            # column P is a formula (SUMPRODUCT of i*(i+1)); if the file was saved without
            # cached values, rebuild it from the prestige level instead of reading zero.
            sac = prestige * (prestige + 1) * (prestige + 2) // 3
        rows.append({
            "id": name_key(display),
            "display": display,
            "bits": total - sac * SAC_UNIT,          # J - P*1,000,000
            "total": total,
            "sacrificed": sac,
            "rank": str(raw[i_rank] if len(raw) > i_rank else "").strip(),
            "title": str(raw[i_title] if len(raw) > i_title else "").strip(),
            "prestige": prestige,
        })
    wb.close()
    return rows


def sync_bank(db, blob=None, origin=None) -> dict:
    if blob is None:
        blob, origin = fetch_bank()
    rows = parse_bank(blob)
    by_id = {r["id"]: r for r in rows}
    touched = 0
    for user in db.query(User).all():
        row = by_id.get(name_key(user.lichess_id))
        if not row:
            continue
        user.bank_bits, user.bank_rank = row["bits"], row["rank"]
        user.bank_total, user.prestige = row["total"], row["prestige"]
        user.bank_synced_at = utcnow()
        user.approved = True          # being in the bank is the proof of membership
        touched += 1
    meta_set(db, "bank_origin", origin or "")
    meta_set(db, "bank_synced_at", utcnow().isoformat())
    meta_set(db, "bank_rows", json.dumps(rows, separators=(",", ":")))
    db.commit()
    return {"rows": len(rows), "users_updated": touched, "origin": origin}


def bank_rows_list(db) -> list:
    blob = meta_get(db, "bank_rows", "")
    return json.loads(blob) if blob else []


def bank_row(db, lichess_id):
    lid = name_key(lichess_id)
    for row in bank_rows_list(db):
        if row["id"] == lid:
            return row
    return None


def apply_bank_to_user(db, user):
    row = bank_row(db, user.lichess_id)
    if not row:
        return False
    user.bank_bits, user.bank_rank = row["bits"], row["rank"]
    user.bank_total, user.prestige = row["total"], row["prestige"]
    user.approved = True
    user.bank_synced_at = utcnow()
    db.commit()
    return True


# --------------------------------------------------------------------------- auth
COOKIE, ADMIN_COOKIE = "bs_session", "bs_admin"
PBKDF2_ROUNDS = 200_000
_DEV_KEY = secrets.token_bytes(32)


def _b64(raw: bytes) -> str:
    return base64.urlsafe_b64encode(raw).rstrip(b"=").decode()


def _unb64(txt: str) -> bytes:
    return base64.urlsafe_b64decode(txt + "=" * (-len(txt) % 4))


def signing_key() -> bytes:
    return SECRET_KEY.encode() if SECRET_KEY else _DEV_KEY


def sign(payload: dict) -> str:
    body = _b64(json.dumps(payload, separators=(",", ":")).encode())
    return body + "." + _b64(hmac.new(signing_key(), body.encode(), hashlib.sha256).digest())


def unsign(token: str):
    if not token or "." not in token:
        return None
    body, mac = token.rsplit(".", 1)
    if not hmac.compare_digest(mac, _b64(hmac.new(signing_key(), body.encode(),
                                                       hashlib.sha256).digest())):
        return None
    try:
        return json.loads(_unb64(body))
    except Exception:                                     # noqa: BLE001
        return None


def hash_password(password: str) -> str:
    salt = os.urandom(16)
    return "pbkdf2$%d$%s$%s" % (PBKDF2_ROUNDS, _b64(salt),
                                _b64(hashlib.pbkdf2_hmac("sha256", password.encode(), salt,
                                                         PBKDF2_ROUNDS)))


def check_password(password: str, stored: str) -> bool:
    try:
        _, rounds, salt, digest = stored.split("$")
        return hmac.compare_digest(_b64(hashlib.pbkdf2_hmac("sha256", password.encode(),
                                                            _unb64(salt), int(rounds))), digest)
    except Exception:                                     # noqa: BLE001
        return False


def pkce():
    verifier = _b64(secrets.token_bytes(48))
    return verifier, _b64(hashlib.sha256(verifier.encode()).digest())


class LichessError(Exception):
    pass


def lichess_url(state, challenge):
    return LICHESS_AUTHORIZE + "?" + urllib.parse.urlencode({
        "response_type": "code", "client_id": LICHESS_CLIENT_ID,
        "redirect_uri": LICHESS_REDIRECT, "scope": "preference:read",
        "state": state, "code_challenge": challenge, "code_challenge_method": "S256"})


def lichess_token(code, verifier):
    """Swap the one-time code for a token at /api/token. PKCE, no secret."""
    data = {"grant_type": "authorization_code", "code": code,
            "code_verifier": verifier, "redirect_uri": LICHESS_REDIRECT,
            "client_id": LICHESS_CLIENT_ID}
    req = urllib.request.Request(
        LICHESS_TOKEN, data=urllib.parse.urlencode(data).encode(),
        headers={"Content-Type": "application/x-www-form-urlencoded",
                 "Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=25) as r:
            return json.loads(r.read())
    except urllib.error.HTTPError as exc:
        detail = ""
        try:
            body = json.loads(exc.read())
            detail = body.get("error_description") or body.get("error") or ""
        except Exception:                                  # noqa: BLE001
            detail = ""
        raise LichessError(detail or "Lichess answered %s" % exc.code) from None
    except Exception as exc:                               # noqa: BLE001
        raise LichessError("could not reach lichess.org (%s)" % exc) from None


def lichess_account(token):
    req = urllib.request.Request("https://lichess.org/api/account",
                                 headers={"Authorization": "Bearer " + token})
    with urllib.request.urlopen(req, timeout=20) as r:
        return json.loads(r.read())

# --------------------------------------------------------------------------- styles
# Three moods, one sheet. body[data-w] swaps the accent, the glow and the wash behind
# the whole page, so Wheel I reads green, Wheel II grey and violet, Wheel III fire.
CSS = """
:root{--bg:#0b0d12;--panel:#141821;--line:#242b39;--ink:#e9edf6;--muted:#93a0b8;
--gold:#f2c14e;--good:#4ade80;--bad:#f87171;--acc:#6cc48a;--acc2:#2f7a52;
--glow:rgba(108,196,138,.45);--wash:#16261c;--wash2:#0b0d12}
body[data-w="0"]{--acc:#f2c14e;--acc2:#b98f2e;--glow:rgba(242,193,78,.4);--wash:#1a2130;
--wash2:#0b0d12}
body[data-w="2"]{--acc:#a99cd8;--acc2:#5c5478;--glow:rgba(169,156,216,.45);
--wash:#1d1a2b;--wash2:#0e0d15}
body[data-w="3"]{--acc:#f0904a;--acc2:#9c3618;--glow:rgba(240,144,74,.5);
--wash:#2b1410;--wash2:#0f0908}
*{box-sizing:border-box}
/* the hidden attribute has to beat .pot{display:grid} and .bigbtn{display:inline-flex} */
[hidden]{display:none!important}
body{margin:0;color:var(--ink);transition:background .4s ease;
background:radial-gradient(1100px 560px at 50% -8%,var(--wash),var(--wash2) 62%) fixed;
font:16px/1.55 -apple-system,BlinkMacSystemFont,'Segoe UI',Inter,sans-serif}
a{color:var(--acc)}
.wrap{max-width:1240px;margin:0 auto;padding:26px 20px 70px}
header.top{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:18px}
.brand{font-family:Georgia,'Times New Roman',serif;font-size:1.32rem}
.brand b{color:var(--acc)}
.spacer{flex:1}
.card{background:linear-gradient(180deg,rgba(255,255,255,.045),rgba(255,255,255,.015));
border:1px solid var(--line);border-radius:18px;padding:22px}
.lab{color:var(--muted);font-size:.76rem;text-transform:uppercase;letter-spacing:.14em}
.pill{display:inline-block;padding:4px 10px;border-radius:999px;border:1px solid var(--line);
font-size:.78rem;color:var(--muted)}
.pill.ok{border-color:rgba(74,222,128,.5);color:#9ae6b4}
.pill.no{border-color:rgba(248,113,113,.5);color:#fca5a5}
.pill.acc{border-color:var(--acc);color:var(--acc)}
button,.btn{cursor:pointer;font:inherit;border-radius:12px;border:1px solid var(--line);
background:#1b2130;color:var(--ink);padding:10px 16px;transition:.15s;text-decoration:none;
display:inline-block}
button:hover:not(:disabled){border-color:var(--acc2);transform:translateY(-1px)}
button:disabled{opacity:.45;cursor:not-allowed}
.primary{background:linear-gradient(180deg,var(--acc),var(--acc2));color:#07130c;font-weight:700;
border:0}
body[data-w="2"] .primary{color:#100e1a}
body[data-w="3"] .primary{color:#1a0a04}
input{font:inherit;background:#0e131c;border:1px solid var(--line);color:var(--ink);
border-radius:10px;padding:10px 12px;width:100%}
canvas{display:block;margin:0 auto;max-width:100%}
table{width:100%;border-collapse:collapse;font-size:.9rem}
th,td{text-align:left;padding:8px 10px;border-bottom:1px solid var(--line)}
th{color:var(--muted);font-weight:600;font-size:.74rem;text-transform:uppercase;
letter-spacing:.1em}
.mono{font-variant-numeric:tabular-nums}
.d-pos{color:#4ade80}.d-neg{color:#f87171}
.hint{color:var(--muted);font-size:.86rem}
.err{color:#fca5a5;background:rgba(248,113,113,.08);border:1px solid rgba(248,113,113,.35);
padding:10px 14px;border-radius:12px;margin:12px 0}
footer{margin-top:32px;color:var(--muted);font-size:.8rem;text-align:center}

/* ---- the three tabs */
.tabs{display:flex;gap:10px;flex-wrap:wrap;margin:0 0 14px}
.tab{display:flex;flex-direction:column;gap:2px;align-items:flex-start;padding:10px 16px;
border-radius:14px;border:1px solid var(--line);background:rgba(255,255,255,.02);
color:var(--muted);min-width:150px;line-height:1.2}
.tab b{font-size:.98rem;color:var(--ink);letter-spacing:.01em}
.tab span{font-size:.74rem}
.tab.on{border-color:var(--acc);background:linear-gradient(180deg,rgba(255,255,255,.07),
rgba(255,255,255,.02));box-shadow:0 0 0 1px var(--acc) inset,0 6px 22px -12px var(--glow)}
.tab.on b{color:var(--acc)}
.tab.locked{cursor:not-allowed;opacity:.55}
.tab.locked b::after{content:' \\1F512';font-size:.8em}
.tab:disabled{transform:none}

/* ---- the glowing progress bar */
.progcard{border:1px solid var(--line);border-radius:16px;padding:14px 18px;margin:0 0 16px;
background:linear-gradient(180deg,rgba(255,255,255,.04),rgba(255,255,255,.01));
position:relative;box-shadow:0 0 24px -18px var(--glow)}
.proghead{display:flex;justify-content:space-between;align-items:baseline;gap:12px;
flex-wrap:wrap;margin-bottom:9px}
.proghead b{font-variant-numeric:tabular-nums;letter-spacing:.02em}
.pbar{height:15px;border-radius:999px;background:rgba(255,255,255,.07);position:relative;
overflow:hidden;box-shadow:inset 0 1px 3px rgba(0,0,0,.55)}
.pbar>i{display:block;height:100%;border-radius:999px;width:0;
background:linear-gradient(90deg,var(--acc2),var(--acc));
box-shadow:0 0 10px var(--glow),0 0 26px var(--glow);transition:width .8s cubic-bezier(.2,.8,.2,1)}
.pbar>i::after{content:'';position:absolute;inset:0;border-radius:999px;
background:linear-gradient(90deg,transparent,rgba(255,255,255,.5),transparent);
animation:sheen 2.8s linear infinite}
@keyframes sheen{0%{transform:translateX(-100%)}100%{transform:translateX(100%)}}
@media(prefers-reduced-motion:reduce){.pbar>i::after{animation:none}}

/* ---- the stage */
.stage{display:grid;grid-template-columns:310px minmax(0,1fr) 330px;gap:18px;align-items:start}
@media(max-width:1080px){.stage{grid-template-columns:minmax(0,1fr) 330px}}
@media(max-width:860px){.stage{grid-template-columns:1fr}}
.wheelcard{text-align:center}
.balance{font-family:Georgia,serif;font-size:2.3rem;line-height:1.05;margin:6px 0 2px}
.wagrow{display:flex;gap:7px;flex-wrap:wrap;margin:10px 0 6px}
.wagrow button{padding:7px 12px;font-weight:600;font-size:.86rem}
.wagrow button.on{border-color:var(--acc);background:rgba(255,255,255,.09);color:var(--acc)}
.btnrow{display:flex;gap:10px;flex-wrap:wrap;align-items:stretch;margin-top:12px}
.bigbtn{display:inline-flex;flex-direction:column;align-items:center;justify-content:center;
gap:2px;padding:11px 22px;border-radius:14px;min-width:150px}
.bigbtn b{font-size:1.2rem;letter-spacing:.05em;line-height:1.15}
.bigbtn span{font-size:.78rem;font-weight:600;opacity:.88}
.bigbtn.free{background:linear-gradient(180deg,#7ee0a8,#38a967);color:#06220f;border:0}
.result{min-height:54px;display:flex;align-items:center;justify-content:center;
font-family:Georgia,serif;font-size:1.4rem;color:var(--acc);text-align:center;margin-top:6px}

/* ---- the odds column */
.odds{max-height:none}
.odds ol{list-style:none;margin:10px 0 0;padding:0}
.odds li{display:grid;grid-template-columns:26px minmax(0,1fr) auto;gap:8px;align-items:center;
padding:5px 0;border-bottom:1px solid rgba(255,255,255,.05);font-size:.88rem}
.odds li em{font-style:normal;width:24px;text-align:center;font-size:1rem}
.odds li b{font-weight:600;font-variant-numeric:tabular-nums;color:var(--ink)}
.odds li s{text-decoration:none;color:var(--muted);font-variant-numeric:tabular-nums;
font-size:.82rem}
.odds li .tag{font-size:.66rem;color:var(--muted);border:1px solid var(--line);
border-radius:999px;padding:1px 6px;margin-left:6px;white-space:nowrap}
.odds li.special b{color:var(--acc)}

/* ---- the jackpot */
.pot{display:grid;grid-template-columns:104px minmax(0,1fr);gap:14px;align-items:center;
border:1px solid rgba(240,144,74,.35);border-radius:16px;padding:14px;
background:radial-gradient(120% 120% at 20% 0%,rgba(240,144,74,.14),rgba(0,0,0,0) 60%)}
.pot .big{font-family:Georgia,serif;font-size:1.72rem;line-height:1.05;
color:#ffd08a;font-variant-numeric:tabular-nums;text-shadow:0 0 18px rgba(240,144,74,.45)}
.pot small{color:var(--muted);display:block;font-size:.78rem;margin-top:3px}
.piggy{width:104px;height:88px}

/* ---- explainer */
details.how{margin-top:14px}
details.how summary{cursor:pointer;color:var(--acc);font-size:.9rem;list-style:none}
details.how summary::-webkit-details-marker{display:none}
details.how summary::before{content:'\\25B8  ';display:inline-block;transition:.2s}
details.how[open] summary::before{transform:rotate(90deg)}
details.how div.hint{margin-top:8px;line-height:1.6}
details.how h4{margin:12px 0 4px;font-size:.86rem;letter-spacing:.02em}
details.how p{margin:0 0 6px;color:var(--muted);font-size:.85rem;line-height:1.6}

/* ---- spin log */
.wtag{display:inline-block;min-width:24px;text-align:center;font-size:.72rem;
border:1px solid var(--line);border-radius:6px;padding:0 4px;color:var(--muted);margin-right:6px}
.wtag.w1{border-color:rgba(108,196,138,.5);color:#9fe0b4}
.wtag.w2{border-color:rgba(169,156,216,.5);color:#c9c0ef}
.wtag.w3{border-color:rgba(240,144,74,.5);color:#f7b98a}
"""

# --------------------------------------------------------------------------- the page
def esc(v):
    return (str(v or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def json_for_script(value) -> str:
    """JSON that is safe to drop inside a <script> block."""
    return json.dumps(value, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")


def page(body, title="Real Humans · Mystery BITS", wheel=1):
    return ("<!doctype html><html lang=en><head><meta charset=utf-8>"
            "<meta name=viewport content='width=device-width,initial-scale=1'>"
            "<title>%s</title><style>%s</style></head><body data-w=\"%d\">"
            "<div class=wrap>%s</div></body></html>"
            % (esc(title), CSS, int(wheel or 1), body))


def nicetime(raw):
    try:
        return _aware(dt.datetime.fromisoformat(raw)).strftime("%d %b %Y, %H:%M UTC")
    except (TypeError, ValueError):
        return ""


def ago(raw):
    """'4m ago', '3h ago', for the last spin column."""
    if not raw:
        return ""
    try:
        when = _aware(dt.datetime.fromisoformat(raw))
    except (TypeError, ValueError):
        return ""
    secs = (utcnow() - when).total_seconds()
    if secs < 60:
        return "just now"
    if secs < 3600:
        return "%dm ago" % int(secs // 60)
    if secs < 86400:
        return "%dh ago" % int(secs // 3600)
    return "%dd ago" % int(secs // 86400)


def wheel_badge(wheel) -> str:
    w = int(wheel or 1)
    return '<span class="wtag w%d">%s</span>' % (w, WHEELS.get(w, WHEELS[1])["numeral"])


WHEEL_JS = r"""
/* ------------------------------------------------------------------ the three wheels
   Everything the page needs arrives in DATA: the slices, their chances, what each one
   hands over, and how far along the bar is. Nothing here decides a result; the server
   picks the slice from a seed it committed to before the spin was paid for. */
const DATA = __DATA__;
const CODES = __CODES__;
const SHEET_URL = '/emoji-sheet.png';

const W = {};
DATA.wheels.forEach(w => { W[w.id] = w; });
let active = DATA.active || 1;
const wagered = {1: 1, 2: 1, 3: 1};

const SIZE = 600, DPR = Math.min(2, window.devicePixelRatio || 1);
const cv = document.getElementById('wheel'), ctx = cv.getContext('2d');
cv.width = SIZE * DPR; cv.height = SIZE * DPR;
cv.style.width = SIZE + 'px'; cv.style.height = SIZE + 'px';
ctx.scale(DPR, DPR);
const C = SIZE / 2, R = C - 24, HUB = R * 0.19;

let angle = -Math.PI / 2, spinning = false, spinId = null, spinWheel = 1;

/* Emoji: use the device's own colour emoji first. If there is none, fall back to the
   sheet that ships with the site, so nobody is ever shown an empty box. */
const sheet = new Image();
let sheetReady = false, useSheet = {};
sheet.onload = () => { sheetReady = true; drawWheel(); };
sheet.src = SHEET_URL;

function emojiIsColour(ch) {
  try {
    const c = document.createElement('canvas');
    c.width = c.height = 32;
    const x = c.getContext('2d');
    x.textAlign = 'center'; x.textBaseline = 'middle'; x.fillStyle = '#000';
    x.font = '26px "Apple Color Emoji","Segoe UI Emoji","Noto Color Emoji",' +
             '"Twemoji Mozilla","EmojiOne Color",sans-serif';
    x.fillText(ch, 16, 17);
    const d = x.getImageData(0, 0, 32, 32).data;
    let ink = 0, chroma = 0;
    for (let i = 0; i < d.length; i += 4) {
      if (d[i + 3] > 60) {
        ink++;
        chroma += Math.max(d[i], d[i + 1], d[i + 2]) - Math.min(d[i], d[i + 1], d[i + 2]);
      }
    }
    return ink > 40 && chroma / Math.max(1, ink) > 10;
  } catch (e) { return false; }
}

function pickSheet() {
  useSheet = {};
  [1, 2, 3].forEach(n => {
    W[n].slices.forEach(s => {
      if (!(s.emoji in useSheet)) useSheet[s.emoji] = !emojiIsColour(s.emoji);
    });
  });
}

/* ---- where the label goes: text spans a band, the emoji sits out at the rim ---- */
function fitLines(lines, maxLen, maxH) {
  let fs = Math.min(19, R * 0.062);
  for (; fs >= 7.5; fs -= 0.4) {
    ctx.font = '700 ' + fs.toFixed(1) + 'px Helvetica,Arial,sans-serif';
    let widest = 0;
    for (let i = 0; i < lines.length; i++) {
      widest = Math.max(widest, ctx.measureText(lines[i]).width);
    }
    if (widest <= maxLen && lines.length * fs * 1.13 <= maxH) return fs;
  }
  return 7.5;
}

function drawWheel() {
  const spec = W[active], S = spec.slices, SEG = S.length;
  const STEP = 2 * Math.PI / SEG;
  ctx.clearRect(0, 0, SIZE, SIZE);

  ctx.save();
  ctx.shadowColor = 'rgba(0,0,0,.55)'; ctx.shadowBlur = 26; ctx.shadowOffsetY = 8;
  ctx.beginPath(); ctx.arc(C, C, R, 0, 2 * Math.PI);
  ctx.fillStyle = '#10131a'; ctx.fill();
  ctx.restore();

  for (let i = 0; i < SEG; i++) {
    const a0 = angle + i * STEP;
    ctx.beginPath(); ctx.moveTo(C, C); ctx.arc(C, C, R, a0, a0 + STEP); ctx.closePath();
    ctx.fillStyle = S[i].color; ctx.fill();
  }
  ctx.strokeStyle = 'rgba(255,255,255,.30)'; ctx.lineWidth = 1.4;
  for (let i = 0; i < SEG; i++) {
    const a = angle + i * STEP;
    ctx.beginPath(); ctx.moveTo(C, C);
    ctx.lineTo(C + Math.cos(a) * R, C + Math.sin(a) * R); ctx.stroke();
  }

  const emojiS = R * 0.082, gap = R * 0.020;
  const outer = R * 0.955, textOuter = outer - emojiS - gap, textInner = R * 0.40;
  const maxLen = textOuter - textInner;
  const maxH = 2 * textInner * Math.sin(STEP / 2) * 0.84;

  for (let i = 0; i < SEG; i++) {
    const mid = angle + (i + 0.5) * STEP;
    ctx.save(); ctx.translate(C, C); ctx.rotate(mid);
    const lines = S[i].lines && S[i].lines.length ? S[i].lines : [S[i].label];
    const fs = fitLines(lines, maxLen, maxH);
    ctx.fillStyle = S[i].ink || '#222';
    ctx.font = '700 ' + fs.toFixed(1) + 'px Helvetica,Arial,sans-serif';
    ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
    const lh = fs * 1.13;
    const top = -lh * (lines.length - 1) / 2;
    for (let k = 0; k < lines.length; k++) {
      ctx.fillText(lines[k], textOuter, top + k * lh);
    }

    const ex = outer - emojiS / 2;
    const idx = CODES.indexOf(S[i].tile || S[i].code);
    if (useSheet[S[i].emoji] && sheetReady && idx >= 0) {
      ctx.drawImage(sheet, idx * 72, 0, 72, 72, ex - emojiS / 2, -emojiS / 2, emojiS, emojiS);
    } else {
      ctx.font = (emojiS * 0.96).toFixed(1) + 'px "Apple Color Emoji","Segoe UI Emoji",' +
                 '"Noto Color Emoji",sans-serif';
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(S[i].emoji, ex, emojiS * 0.05);
    }
    ctx.restore();
  }

  ctx.beginPath(); ctx.arc(C, C, HUB, 0, 2 * Math.PI);
  ctx.fillStyle = '#0e1118'; ctx.fill();
  ctx.lineWidth = 2; ctx.strokeStyle = 'rgba(255,255,255,.18)'; ctx.stroke();
  ctx.fillStyle = spinning ? 'rgba(255,255,255,.35)' : getComputedStyle(document.body).getPropertyValue('--acc');
  ctx.font = '700 ' + Math.round(HUB * 0.42) + 'px Helvetica,Arial,sans-serif';
  ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
  ctx.fillText(spinning ? '\u2026' : spec.numeral, C, C + 1);

  const tipX = C + R * 0.955, baseX = C + R * 1.085, hh = R * 0.088;
  ctx.beginPath();
  ctx.moveTo(tipX, C); ctx.lineTo(baseX, C - hh); ctx.lineTo(baseX, C + hh);
  ctx.closePath();
  ctx.fillStyle = '#ffffff'; ctx.fill();
  ctx.lineWidth = 2; ctx.strokeStyle = 'rgba(120,128,148,.9)'; ctx.stroke();
}

/* ------------------------------------------------------------------ the panels */
function money(n) { return Number(n || 0).toLocaleString(); }

function renderTabs() {
  const box = document.getElementById('tabs');
  box.innerHTML = DATA.wheels.map(w => {
    const cls = 'tab' + (w.id === active ? ' on' : '') + (w.unlocked ? '' : ' locked');
    let sub = w.unlocked ? ('\u2713 ' + (w.id === 3 ? '30,000 a spin' : money(w.cost) + ' a spin'))
      : (w.progress.have >= w.progress.need ? 'open' :
         money(w.progress.have) + ' / ' + money(w.progress.need));
    if (w.free && w.free.count) sub += ' \u00b7 ' + w.free.count + ' free';
    return '<button type=button class="' + cls + '" data-w="' + w.id + '"' +
      (w.unlocked || w.id === active ? '' : ' aria-disabled=true') + '>' +
      '<b>' + w.numeral + ' \u00b7 ' + w.name + '</b><span>' + sub + '</span></button>';
  }).join('');
  box.querySelectorAll('button').forEach(b => {
    b.onclick = () => chooseWheel(Number(b.dataset.w));
  });
}

function renderProgress() {
  const w = W[active], p = w.progress;
  const pct = Math.max(0, Math.min(100, p.need ? (p.have / p.need) * 100 : 0));
  const done = p.have >= p.need;
  document.getElementById('progLab').textContent = p.secret
    ? 'The end of this bar is not written down anywhere'
    : ('Progress to Wheel ' + (active === 1 ? 'II' : 'III'));
  document.getElementById('progVal').innerHTML = p.secret
    ? (money(p.have) + ' / ' + money(p.need) + ' BITS won here')
    : (money(p.have) + ' / ' + money(p.need) + ' spins' + (done ? ' \u2014 open' : ''));
  const bar = document.getElementById('progBar');
  bar.style.width = pct.toFixed(2) + '%';
  bar.parentNode.setAttribute('aria-valuenow', Math.round(pct));
  document.getElementById('progNote').textContent = p.secret
    ? (done ? 'Somebody has reached the end. That is all that is known.'
            : 'Nobody who has reached it has said what waits there. It is a special title, ' +
              'and something else.')
    : (done ? 'You can spin this wheel.' :
       'Spins you left before the wheel stopped do not count towards this.');
}

function renderOdds() {
  const w = W[active];
  document.getElementById('oddsTitle').textContent = 'Chances on Wheel ' + w.numeral;
  document.getElementById('oddsList').innerHTML = w.slices.map(s => {
    const tag = s.kind === 'free' ? '<span class=tag>that wheel</span>'
      : s.kind === 'item' ? (s.note ? '<span class=tag>' + s.note + '</span>' : '')
      : s.kind === 'jackpot' ? '<span class=tag>the whole pot</span>' : '';
    const tip = s.kind === 'jackpot'
      ? 'Every paid spin on Wheel III drops its 30,000 BITS into one pot. Whoever lands here ' +
        'takes the whole pot, and it starts again from zero.'
      : s.kind === 'item' && s.label.indexOf('fragment') >= 0
      ? 'Three fragments of the same kind become one Booster of that kind. Fragments are ' +
        'yours for good; the Booster is handed to you by the admin.'
      : s.label;
    return '<li class="' + (s.kind !== 'bits' ? 'special' : '') + '" title="' + tip + '">' +
      '<em>' + s.emoji + '</em><b>' + s.label + tag + '</b><s>' +
      s.pct.toFixed(2) + '%</s></li>';
  }).join('');
  const avg = w.average;
  document.getElementById('oddsFoot').textContent =
    'A spin costs ' + money(w.cost) + ' BITS' +
    (w.wager ? ' (or a multiple of it)' : '') + '. Over a great many spins the wheel hands ' +
    'back about ' + money(Math.round(avg.per)) + ' BITS a spin, which is ' + avg.pct +
    '% of the price. Items, tokens and free spins are not BITS and are not counted in that.';
}

function renderHud() {
  const w = W[active];
  document.getElementById('wtag').textContent = 'Wheel ' + w.numeral + ' \u00b7 ' + w.name;
  document.getElementById('wtag').className = 'pill acc';
  document.getElementById('wtagline').textContent = w.tag;
  document.getElementById('moodline').textContent = w.mood;

  const wagBox = document.getElementById('wagBox');
  if (!w.wager) {
    wagBox.innerHTML = '<div class="hint">No wager here. Every spin is ' + money(w.cost) +
      ' BITS, and the whole price goes into the jackpot.</div>';
  } else if (!DATA.multi) {
    wagBox.innerHTML = '<div class="hint">Below Sovereign the wager is fixed at one, so a ' +
      'spin is ' + money(w.cost) + ' BITS.</div>';
  } else {
    const steps = w.id === 1 ? [1, 2, 5, 10, 25, 50, 100, 250, 500, 1000]
                             : [1, 2, 3, 5, 10];
    wagBox.innerHTML = '<div class="lab">Wager</div><div class=wagrow>' +
      steps.filter(s => s <= w.maxWager).map(s =>
        '<button type=button data-w2="' + s + '"' +
        (wagered[w.id] === s ? ' class=on' : '') + '>x' + money(s) + '</button>').join('') +
      '</div><input id=wager type=number min=1 step=1 max=' + w.maxWager +
      ' value=' + wagered[w.id] + ' style="max-width:150px">';
    wagBox.querySelectorAll('button[data-w2]').forEach(b => {
      b.onclick = () => { setWager(w.id, Number(b.dataset.w2)); };
    });
    document.getElementById('wager').oninput = e => {
      let v = parseInt(e.target.value || '1', 10);
      if (!(v > 0)) v = 1;
      wagered[w.id] = Math.min(w.maxWager, v);
      setCost();
    };
  }
  setCost();
  renderFree();
  document.getElementById('go').textContent = '';
  document.getElementById('go').innerHTML = '<b>SPIN</b><span id=costbtn>' +
    money(w.cost * wagered[w.id]) + ' BITS</span>';
}

function setWager(wid, wager) {
  wagered[wid] = Math.max(1, Math.min(W[wid].maxWager, wager));
  renderHud();
}

function setCost() {
  const w = W[active], n = w.cost * (w.wager ? wagered[w.id] : 1);
  const btn = document.getElementById('costbtn');
  if (btn) btn.textContent = money(n) + ' BITS';
  document.getElementById('costline').innerHTML = w.wager
    ? ('This spin: <b>' + money(n) + '</b> BITS (\u00d7' + money(wagered[w.id]) + ')')
    : ('This spin: <b>' + money(n) + '</b> BITS');
}

function renderFree() {
  const w = W[active], btn = document.getElementById('gofree'),
    label = document.getElementById('freebtn'), hint = document.getElementById('freehint');
  const fs = w.free;
  if (!fs || !fs.count) {
    btn.hidden = true; hint.hidden = true; return;
  }
  btn.hidden = false; btn.disabled = spinning;
  hint.hidden = false;
  let txt = '\u00d7' + fs.count + ' left at wager \u00d7' + money(fs.wager);
  const rest = (fs.queue || []).slice(1);
  if (rest.length) txt += ' \u00b7 then ' + rest.map(q =>
    '\u00d7' + q.count + ' at \u00d7' + money(q.wager)).join(', ');
  label.textContent = txt;
}

function renderPot() {
  const box = document.getElementById('potCard');
  if (active !== 3) { box.hidden = true; return; }
  box.hidden = false;
  document.getElementById('potVal').textContent = money(DATA.jackpot) + ' BITS';
}

function busy(on) {
  spinning = on;
  document.getElementById('go').disabled = on;
  document.getElementById('gofree').disabled = on;
  document.querySelectorAll('.wagrow button').forEach(b => { b.disabled = on; });
  const inp = document.getElementById('wager');
  if (inp) inp.disabled = on;
  drawWheel();
}

function show(msg, bad) {
  const el = document.getElementById('err');
  el.textContent = msg || '';
  el.style.display = msg ? 'block' : 'none';
  el.style.color = bad === false ? '#9ae6b4' : '#fca5a5';
  el.style.background = bad === false ? 'rgba(74,222,128,.08)' : 'rgba(248,113,113,.08)';
  el.style.borderColor = bad === false ? 'rgba(74,222,128,.35)' : 'rgba(248,113,113,.35)';
}

function chooseWheel(id) {
  if (spinning) return;
  const w = W[id];
  if (!w.unlocked) {
    show('Wheel ' + w.numeral + ' is locked. ' + money(w.progress.have) + ' of ' +
      money(w.progress.need) + ' spins on Wheel ' +
      (id === 2 ? 'I' : 'II') + ' so far.', true);
    return;
  }
  show('');
  active = id;
  document.body.setAttribute('data-w', String(id));
  renderTabs(); renderProgress(); renderHud(); renderOdds(); renderPot();
  drawWheel();
}

/* ------------------------------------------------------------------ spinning */
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
const VMAX = 0.0125, RAMP = 480, GUARD = 16000;

async function askReveal() {
  let last = 'The spin was refused.';
  for (let i = 0; i < 8; i++) {
    const r = await fetch('/api/spin/' + spinId + '/reveal', {method: 'POST'});
    const d = await r.json();
    if (r.ok) return d;
    last = d.detail || last;
    if (!/still turning/i.test(last)) throw new Error(last);
    await sleep(300);
  }
  throw new Error(last);
}

async function collect() {
  let last = 'The spin could not be collected.';
  for (let i = 0; i < 4; i++) {
    const r = await fetch('/api/spin/' + spinId + '/claim', {method: 'POST'});
    const d = await r.json();
    if (r.ok) return d;
    last = d.detail || last;
    await sleep(400);
  }
  throw new Error(last);
}

function refresh() {
  return fetch('/api/me', {cache: 'no-store'}).then(r => r.ok ? r.json() : null)
    .then(d => {
      if (!d) return null;
      DATA.wheels.forEach(w => {
        const fresh = d.wheels.filter(x => x.id === w.id)[0];
        if (!fresh) return;
        w.unlocked = fresh.unlocked; w.progress = fresh.progress; w.free = fresh.free;
      });
      DATA.jackpot = d.jackpot;
      renderTabs(); renderProgress(); renderHud(); renderOdds(); renderPot();
      return d;
    });
}

async function spin(useFree) {
  if (spinning) return;
  show('');
  document.getElementById('result').textContent = '';
  const spec = W[active];
  const body = useFree ? {free: true, wheel: active}
                       : {wheel: active, wager: spec.wager ? wagered[active] : 1};
  const res = await fetch('/api/spin', {
    method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(body)
  });
  const j = await res.json();
  if (!res.ok) { show(j.detail || 'Refused', true); return; }
  if (j.capped) {
    show('Below Sovereign the wager is fixed at one, so this spin cost ' + money(j.cost) +
      ' BITS.', false);
  }
  spinId = j.spin_id; spinWheel = active;
  document.getElementById('balance').textContent = money(j.balance);
  document.getElementById('commit').textContent =
    'committed before you paid: ' + j.commit.slice(0, 16) + '\u2026';
  busy(true);

  const revealMs = Number(j.reveal_after || DATA.reveal) * 1000;
  let last = performance.now(), t0 = last, speed = 0, armed = false, got = null, bad = null;
  await new Promise(done => (function frame(now) {
    const dt = Math.min(50, now - last); last = now;
    const t = now - t0;
    speed = t < RAMP ? VMAX * (t / RAMP) * (t / RAMP) : VMAX;
    angle += speed * dt;
    drawWheel();
    if (!armed && t >= revealMs) {
      armed = true;
      askReveal().then(d => { got = d; }).catch(e => { bad = e; });
    }
    if (bad || got || t > GUARD) return done();
    requestAnimationFrame(frame);
  })(last));

  if (bad) { busy(false); show(bad.message, true); loadLog(); return; }
  if (!got) {
    busy(false);
    show('The server stopped answering. Nothing was collected. Check Your spins.', true);
    loadLog(); return;
  }

  const d = got;
  const TAU = 2 * Math.PI, SEG = W[spinWheel].slices.length, STEP = TAU / SEG;
  const wrap = a => a - TAU * Math.floor(a / TAU);
  const target = -(d.index + 0.5) * STEP;
  const from = angle;
  const delta = wrap(target - from) + TAU;
  const T = 3 * delta / Math.max(speed, VMAX * 0.7);
  const startedAt = performance.now();
  await new Promise(done => (function ease(now) {
    const p = Math.min(1, (now - startedAt) / T);
    angle = from + delta * (1 - Math.pow(1 - p, 3));
    drawWheel();
    p < 1 ? requestAnimationFrame(ease) : done();
  })(startedAt));
  angle = wrap(from + delta);
  drawWheel();

  let paid;
  try {
    paid = await collect();
  } catch (e) {
    busy(false);
    document.getElementById('result').textContent = '\u{1F621} Lost: the wheel was not collected';
    show(e.message || 'The spin was not collected.', true);
    document.getElementById('commit').textContent = '';
    loadLog(); refresh();
    return;
  }

  const slice = W[spinWheel].slices[paid.index];
  const em = slice ? slice.emoji : '';
  let line;
  if (paid.kind === 'jackpot') {
    line = em + ' JACKPOT  \u2192  +' + money(paid.paid) + ' BITS';
  } else if (paid.kind === 'free') {
    line = em + ' ' + paid.label + '  \u2192  +' + money(paid.prize_qty) + ' free spins';
  } else if (paid.kind === 'item') {
    line = em + ' ' + paid.label + ' \u00d7 ' + money(paid.prize_qty);
    if (paid.note) line += '  (' + paid.note + ')';
  } else {
    line = em + ' ' + paid.label + '  \u2192  +' + money(paid.paid) + ' BITS';
  }
  document.getElementById('result').textContent = line;
  document.getElementById('balance').textContent = money(paid.balance);
  document.getElementById('commit').textContent =
    'wager \u00d7' + money(paid.wager) + ' \u00b7 seed ' + paid.seed.slice(0, 12) +
    '\u2026 \u00b7 sha256 ' + paid.commit.slice(0, 12) + '\u2026';
  busy(false);
  await refresh();
  loadLog();
}

async function loadLog() {
  const rows = await (await fetch('/api/spins')).json();
  document.querySelector('#log tbody').innerHTML = rows.map(s => {
    const slice = (W[s.wheel] || W[1]).slices[s.index];
    let what;
    if (s.status === 'settled') {
      if (s.kind === 'jackpot') what = s.label + '  +' + money(s.payout);
      else if (s.kind === 'item') what = s.label + ' \u00d7 ' + money(s.prize_qty) +
        (s.note ? ' (' + s.note + ')' : '');
      else if (s.kind === 'free') what = s.label + '  +' + money(s.prize_qty) + ' free';
      else what = s.label + '  +' + money(s.payout);
      if (slice) what = slice.emoji + ' ' + what;
    } else if (s.status === 'forfeited') {
      what = '\u{1F621} left before it stopped';
    } else {
      what = 'still turning\u2026';
    }
    return '<tr><td class=mono>' + new Date(s.created_at).toLocaleString() + '</td>' +
      '<td>' + s.badge + '</td>' +
      '<td class=mono>x' + money(s.wager) + (s.free ? ' free' : '') + '</td>' +
      '<td>' + what + '</td>' +
      '<td class=mono>' + (s.cost ? '-' + money(s.cost) : 'free') + '</td></tr>';
  }).join('') || '<tr><td colspan=5 class=hint>No spins yet.</td></tr>';
}

/* ------------------------------------------------------------------ wiring */
pickSheet();
document.getElementById('go').onclick = () => spin(false);
document.getElementById('gofree').onclick = () => spin(true);
cv.onclick = () => spin(false);
cv.style.cursor = 'pointer';
renderTabs(); renderProgress(); renderHud(); renderOdds(); renderPot();
drawWheel();
loadLog();

/* Tell the admin page this window is still open. Only while the tab is visible, so a
   forgotten tab does not look like somebody sitting in front of it. */
const PING_MS = Number(DATA.ping) * 1000;
function ping() {
  if (document.visibilityState === 'hidden') return;
  fetch('/api/ping', {cache: 'no-store'}).catch(() => {});
}
ping();
setInterval(ping, PING_MS);
document.addEventListener('visibilitychange', ping);
setInterval(() => {                                   /* the pot moves while you watch */
  if (spinning || active !== 3) return;
  fetch('/api/me', {cache: 'no-store'}).then(r => r.ok ? r.json() : null).then(d => {
    if (!d) return;
    DATA.jackpot = d.jackpot;
    document.getElementById('potVal').textContent = money(d.jackpot) + ' BITS';
  }).catch(() => {});
}, 20000);
"""

# --------------------------------------------------------------------------- player page
def delta_line(delta: int) -> str:
    """The sentence under the balance: where that number comes from, change coloured."""
    if delta == 0:
        return ("The BITS above is your BITS Bank figure, unchanged: you have not spun since "
                "the admin last updated.")
    word, cls = ("plus", "d-pos") if delta > 0 else ("minus", "d-neg")
    return ('The BITS above is your BITS Bank figure <b class="%s">%s %s</b>: everything you '
            'have done here since the admin last updated.' % (cls, word, f"{abs(delta):,}"))


PIGGY = """
<svg class=piggy viewBox="0 0 104 88" aria-hidden=true>
  <ellipse cx=48 cy=52 rx=36 ry=27 fill="#e08a52"/>
  <ellipse cx=48 cy=52 rx=36 ry=27 fill=none stroke="#8d4a20" stroke-width=2/>
  <circle cx=30 cy=48 r=2.8 fill="#5d2a10"/>
  <ellipse cx=84 cy=54 rx=9 ry=8 fill="#f2a571" stroke="#8d4a20" stroke-width=2/>
  <circle cx=81 cy=52 r=1.5 fill="#6d3218"/><circle cx=86 cy=55 r=1.5 fill="#6d3218"/>
  <path d="M26 27 l11 -9 l4 13 z" fill="#e08a52" stroke="#8d4a20" stroke-width=2/>
  <rect x=24 y=73 width=11 height=13 rx=3 fill="#d07a44" stroke="#8d4a20" stroke-width=1.5/>
  <rect x=58 y=73 width=11 height=13 rx=3 fill="#d07a44" stroke="#8d4a20" stroke-width=1.5/>
  <rect x=38 y=23 width=22 height=3.4 rx=1.7 fill="#5d2a10"/>
  <circle cx=49 cy=12 r=6.4 fill="#ffcf6b" stroke="#a86b1e" stroke-width=1.6/>
</svg>"""

HOW = """
<details class=how><summary>How these wheels work</summary>
<div class=hint>
<h4>Equal slices, unequal chances</h4>
<p>The wheel is drawn with every slice the same size, because that is what a wheel looks like.
The list above is the real chance of each slice, and those numbers do not match the sizes.
Each wheel changes what a spin costs, what it can give and who may spin it.</p>

<h4>Unlocking</h4>
<p>Wheel I is open to everybody. Wheel II opens after 100 finished spins on Wheel I.
Wheel III opens after 1,000 finished spins on Wheel II. The wager makes no difference, only
the number of spins, and a spin only counts once it has been collected. If you close the page
before the wheel stops, that spin is gone and it does not count.</p>
<p>You can always go back to a lower wheel. The bar belongs to the wheel you are on and does
not change when you switch tabs.</p>

<h4>The Progressive Legacy Jackpot</h4>
<p>Every paid spin on Wheel III puts its whole price, 30,000 BITS, into one pot that all
players share. It is not taken from you as a prize and not paid to you as a win; it goes in
the pot. When somebody lands on the jackpot slice, that player is paid the whole pot and the
pot starts again from zero. Free spins do not feed it. A spin left before the wheel stops
keeps its 30,000 in the pot.</p>

<h4>Fragments and boosters</h4>
<p>Wheel II can give a fragment: x2, x3, x5 or x10. Three fragments of the same kind quietly
become one Booster of that kind, and the admin hands the Booster to you. Your fragments are
kept for good and are never cleared, so nothing is wasted if you stop at two. Wheel III gives
finished boosters directly, including the x20 BITS Booster, which cannot be bought anywhere.</p>

<h4>What a wager multiplies</h4>
<p>On Wheel I and Wheel II the wager multiplies BITS prizes and the three logo prizes. It
never multiplies fragments, boosters, tokens or free spins, because those are counted in
things rather than in BITS. Free spins keep the wager they were won at and belong to the
wheel that paid them out, so a Wheel II free spin is spent on Wheel II.</p>

<h4>Checking a spin</h4>
<p>Before you pay, the server commits to a secret seed and shows you its SHA-256. When the
wheel stops it shows the seed. The slice is the seed read as a number, taken modulo 10,000,
and looked up in the weights. Anybody can recompute it.</p>
</div></details>"""


def index_page(user, payload, delta):
    js = (WHEEL_JS
          .replace("__DATA__", json_for_script(payload))
          .replace("__CODES__", json_for_script(EMOJI_CODES)))
    admin_link = ('<a class=btn href="/admin">Admin</a>' if is_admin(user) else "")
    rank = user.bank_rank or ""
    multi = payload["multi"]
    rank_pill = ('<span class="pill ok">%s \u00b7 multiples unlocked</span>' % esc(rank)
                 if multi else '<span class="pill no">%s \u00b7 wager fixed at one</span>'
                 % esc(rank or "no rank"))
    zero = ""
    if payload["available"] <= 0:
        zero = ('<div class=err><b>You have no BITS yet.</b> Everything is visible, but a spin '
                'costs BITS, so the buttons will refuse until the bank has a figure for you. '
                '<a href="%s" target=_blank rel=noopener>How BITS work</a>. If your Lichess name '
                'is not in the BITS Bank yet, say so in the team chat and it gets added.</div>'
                % esc(BITS_HELP))
    return page("""
<header class=top><div class=brand>Real Humans \u00b7 <b>Mystery BITS</b></div>
<div class=spacer></div><span class=pill>@%s</span>%s%s
<a class=btn href="/logout">Log out</a></header>
%s
<div class=tabs id=tabs></div>
<div class=progcard>
  <div class=proghead><span class=hint id=progLab></span><b id=progVal></b></div>
  <div class=pbar role=progressbar aria-valuemin=0 aria-valuemax=100 aria-valuenow=0>
    <i id=progBar></i></div>
  <div class=hint id=progNote style="margin-top:8px"></div>
</div>
<div class=stage>
  <div class=card>
    <div class=hint id=wtagline style="min-height:2.6em"></div>
    <div class=lab style="margin-top:6px">Your BITS</div>
    <div class=balance id=balance>%s</div>
    <div class=hint>%s</div>
    <hr style="border:0;border-top:1px solid var(--line);margin:16px 0">
    <div id=wagBox></div>
    <div class=hint id=costline style="margin-top:8px"></div>
    <div class=btnrow>
      <button id=go class="bigbtn primary"><b>SPIN</b><span id=costbtn></span></button>
      <button id=gofree class="bigbtn free" hidden><b>FREE SPIN</b><span id=freebtn></span></button>
    </div>
    <div class=hint id=freehint hidden style="margin-top:8px">A free spin costs nothing, keeps
    the wager it was won at, and only exists on the wheel that gave it to you.</div>
    <div id=result class=result></div>
    <div id=err class=err style="display:none"></div>
    <div class=hint id=commit style="word-break:break-all"></div>
    <p class=hint>If you reload or close the page before the wheel stops, the spin is lost:
    the price and anything it landed on. The server enforces that, not this page.</p>
  </div>
  <div class="card wheelcard">
    <div style="margin-bottom:6px"><span class=pill id=wtag></span></div>
    <canvas id=wheel></canvas>
    <div class=hint id=moodline style="margin-top:10px;font-style:italic"></div>
    <div class=pot id=potCard hidden style="margin-top:14px;text-align:left">
      %s
      <div>
        <div class=lab>Progressive Legacy Jackpot</div>
        <div class=big id=potVal>0 BITS</div>
        <small>Every paid spin here adds 30,000 BITS to one pot shared by everyone. Whoever
        lands on the jackpot takes all of it.</small>
      </div>
    </div>
  </div>
  <div class="card odds">
    <div class=lab id=oddsTitle></div>
    <ol id=oddsList></ol>
    <div class=hint id=oddsFoot style="margin-top:12px"></div>
    %s
  </div>
</div>
<div class=card style="margin-top:20px"><div class=lab>Your spins</div>
  <table id=log><thead><tr><th>When</th><th>Wheel</th><th>Wager</th><th>Landed</th>
  <th>Cost</th></tr></thead><tbody></tbody></table>
  <p class=hint style="margin-bottom:0">The last 200 spins on all three wheels.</p></div>
<footer>Real Humans \u00b7 Mystery BITS \u00b7 emoji artwork from Twemoji (CC-BY 4.0) \u00b7
  <a href="%s" target=_blank rel=noopener>BITS guide</a></footer>
<script>%s</script>
""" % (esc(user.display or user.lichess_id), rank_pill, admin_link, zero,
       f"{payload['available']:,}", delta_line(delta), PIGGY, HOW, esc(BITS_HELP), js))


def login_page(message="", notice="", wheel=1):
    local_card = """
  <div class=card><div class=lab>Or a local account</div>
    <p class=hint>Use this only if the Lichess button will not work for you. Your name is
    checked against the bank before you can spin, and a local account can never be admin.</p>
    <form method=post action="/auth/local" style="display:grid;gap:10px">
      <input name=username placeholder="Lichess username, no @" required>
      <input name=password type=password placeholder="Password, 8 characters or more"
        required minlength=8>
      <div style="display:flex;gap:8px">
        <button class=primary name=mode value=login>Log in</button>
        <button name=mode value=register>Register</button></div>
    </form></div>""" if ALLOW_LOCAL else """
  <div class=card><div class=lab>Three wheels</div>
    <p class=hint>Wheel I costs 300 BITS a spin. Wheel II costs 3,000 and opens after 100
    finished spins on Wheel I. Wheel III costs 30,000, has no wager, and opens after 1,000
    finished spins on Wheel II.</p>
    <p class=hint>Your BITS come from the Real Humans BITS Bank. The figure a player sees is
    what the bank said last time the admin updated it, plus or minus everything they have
    done since.</p></div>"""
    top = ('<div class=err>%s</div>' % esc(message)) if message else ""
    low = ('<div class=hint>%s</div>' % esc(notice)) if notice else ""
    return page("""
<header class=top><div class=brand>Real Humans \u00b7 <b>Mystery BITS</b></div></header>
%s
<div class=stage>
  <div class=card><div class=lab>Sign in</div>
    <h2 style="font-family:Georgia,serif;margin:8px 0 6px">Log in with Lichess</h2>
    <p class=hint>One click. This site never sees your Lichess password; Lichess only tells it
    your name.</p>
    <p><a class=btn href="/auth/lichess">Continue with Lichess \u2192</a></p>%s</div>
  <div class=card></div>
  %s
</div>
<footer>Spending BITS are tentative until the admin updates the BITS Bank.</footer>
""" % (top, low, local_card), wheel=wheel)


def state_of(db, user):
    delta = delta_for(db, user.id, user.applied_through or 0)
    return {"bank": int(user.bank_bits or 0), "delta": delta,
            "available": int(user.bank_bits or 0) + delta,
            "multiple": can_wager_multiple(user.bank_rank),
            "rank": user.bank_rank or "", "rank_name": rank_of(user.bank_rank),
            "free": free_state_all(db, user), "jackpot": jackpot_value(db)}

# --------------------------------------------------------------------------- admin
def emoji_of(label):
    for wheel in (WHEEL_1, WHEEL_2, WHEEL_3):
        for slice_ in wheel:
            if slice_["label"] == label:
                return slice_["emoji"]
    return ""


def item_name(key):
    return (ITEMS.get(key) or {}).get("name", key)


def item_emoji(key):
    return (ITEMS.get(key) or {}).get("emoji", "")


def last_spin_of(db):
    """The newest spin of every player, in one round trip."""
    latest = (select(Spin.user_id, func.max(Spin.created_at).label("when_"))
              .group_by(Spin.user_id).subquery())
    q = (select(Spin).join(latest, and_(Spin.user_id == latest.c.user_id,
                                        Spin.created_at == latest.c.when_)))
    return {s.user_id: s for s in db.execute(q).scalars()}


def admin_rows(db):
    bank = {row["id"]: row for row in bank_rows_list(db)}
    left = dict(db.execute(select(Spin.user_id, func.count(Spin.id))
                           .where(Spin.status == "forfeited")
                           .group_by(Spin.user_id)).all())
    newest = last_spin_of(db)
    now = utcnow()
    out = []
    for u in db.query(User).order_by(User.id).all():
        delta = delta_for(db, u.id, u.applied_through or 0)
        cnt, staked = db.execute(select(func.count(Spin.id), func.coalesce(func.sum(Spin.cost), 0))
                                 .where(Spin.user_id == u.id)).one()
        counts = wheel_counts(db, u.id)
        have = unlocked_wheels(db, u, counts)
        spin = newest.get(u.id)
        row = bank.get(name_key(u.lichess_id)) or {}
        out.append({"id": u.id, "display": u.display or u.lichess_id, "rank": u.bank_rank or "",
                    "title": row.get("title", "") or "",
                    "rank_ord": rank_ordinal(u.bank_rank, row.get("title", "")),
                    "bank": int(u.bank_bits or 0), "delta": delta,
                    "available": int(u.bank_bits or 0) + delta, "spins": int(cnt or 0),
                    "staked": int(staked or 0),
                    "counts": counts, "wheels": sorted(have),
                    "in_bank": bool(row), "forfeits": int(left.get(u.id, 0)),
                    "last_wheel": int(spin.wheel or 1) if spin else 0,
                    "last_label": (spin.label or "") if spin else "",
                    "last_emoji": emoji_of(spin.label) if spin else "",
                    "last_status": (spin.status or "") if spin else "",
                    "last_at": _aware(spin.created_at).isoformat() if spin and spin.created_at else "",
                    "last_seen": seen_iso(u), "active": is_active(u, now)})
    return out


def prizes_since(db):
    """Prizes are an obligation only until they have been handed out. 'Mark applied'
    stamps the moment, and that stamp is what empties the list. Fragments are never on
    that list, so pressing the button cannot take one away."""
    raw = meta_get(db, "prizes_applied_at", "")
    try:
        return _aware(dt.datetime.fromisoformat(raw)) if raw else None
    except ValueError:
        return None


def award_totals(db, user_id=None, since=None) -> dict:
    """key -> count, summed over awards that are still owed."""
    q = select(Award.user_id, Award.key, func.sum(Award.qty)).where(Award.qty > 0)
    if user_id is not None:
        q = q.where(Award.user_id == user_id)
    if since is not None:
        q = q.where(Award.created_at > since)
    q = q.group_by(Award.user_id, Award.key)
    out = {}
    for uid, key, qty in db.execute(q):
        out.setdefault(uid, {})[key] = out.setdefault(uid, {}).get(key, 0) + int(qty or 0)
    return out


def prize_rows(db):
    """Everything the admin still has to hand out, per player and per item."""
    since = prizes_since(db)
    totals = award_totals(db, since=since)
    if not totals:
        return []
    names = {u.id: (u.display or u.lichess_id) for u in db.query(User).all()}
    out = []
    for uid, items in totals.items():
        for key, qty in items.items():
            if qty <= 0:
                continue
            out.append({"display": names.get(uid, "?"), "key": key, "name": item_name(key),
                        "emoji": item_emoji(key), "qty": qty})
    out.sort(key=lambda r: (name_key(r["display"]), PRIZE_KEYS.index(r["key"])
                            if r["key"] in PRIZE_KEYS else 99))
    return out


def admin_guard(request, db):
    user = current_user(request, db)
    if is_admin(user):
        return user
    if ADMIN_KEY and hmac.compare_digest(request.headers.get("x-admin-key", ""), ADMIN_KEY):
        return user
    return None


PLAYER_COLS = [("Player", "text"), ("Rank", "num"), ("Bank", "num"), ("Delta", "num"),
               ("Available", "num"), ("Spins", "num"), ("Last Spin", "time"),
               ("Account", "num")]

PRIZE_COLS = [("Player", "text"), ("Prize", "text"), ("Owed", "num")]

SORT_JS = r"""
/* ---- sorting, for any table with data-col / data-type headings. First click puts the
   biggest number on top; the second click turns it round; a third starts again. */
function makeSortable(tbl, storageKey) {
  if (!tbl) return;
  var body = tbl.tBodies[0];
  var order = {col: -1, asc: false};
  function value(tr, col, type) {
    var cell = tr.cells[col];
    var raw = cell ? cell.getAttribute('data-v') : '';
    if (raw === null) raw = '';
    return (type === 'num') ? Number(raw || 0) : String(raw);
  }
  function apply(th, asc) {
    var col = Number(th.getAttribute('data-col'));
    var type = th.getAttribute('data-type') || 'text';
    var rows = Array.prototype.slice.call(body.rows).filter(function(r){ return !r.dataset.fixed; });
    rows.sort(function(a, b) {
      var x = value(a, col, type), y = value(b, col, type);
      var cmp = (type === 'num') ? (x - y)
        : String(x).localeCompare(String(y), 'en', {numeric: true, sensitivity: 'base'});
      if (cmp === 0) return 0;
      return asc ? cmp : -cmp;
    });
    rows.forEach(function(r) { body.appendChild(r); });
    tbl.querySelectorAll('th.sortable').forEach(function(o) {
      var on = (o === th);
      o.classList.toggle('on', on);
      var arrow = o.querySelector('.arrow');
      if (arrow) arrow.textContent = on ? (asc ? '\u25B2' : '\u25BC') : '\u21C5';
    });
    order = {col: col, asc: asc};
    try { localStorage.setItem(storageKey, JSON.stringify(order)); } catch (e) {}
  }
  tbl.querySelectorAll('th.sortable').forEach(function(th) {
    th.addEventListener('click', function() {
      var col = Number(th.getAttribute('data-col'));
      apply(th, (order.col === col) ? !order.asc : false);
    });
  });
  try {
    var saved = JSON.parse(localStorage.getItem(storageKey) || 'null');
    if (saved && saved.col >= 0) {
      var th = tbl.querySelector('th.sortable[data-col="' + saved.col + '"]');
      if (th) apply(th, !!saved.asc);
    }
  } catch (e) {}
}
"""


def admin_page(rows, prizes, meta):
    def row_html(r):
        last = ""
        if r.get("last_at"):
            what = (r["last_emoji"] + " " + r["last_label"]).strip() or r["last_status"]
            last = ('<span class=lastspin>%s%s<span class=ago title="%s">%s</span></span>'
                    % (wheel_badge(r["last_wheel"]), esc(what), esc(nicetime(r["last_at"])),
                       esc(ago(r["last_at"]))))
        else:
            last = '<span class=hint>never</span>'
        seat = "".join(
            '<span class="wtab%s" title="Wheel %s">%s</span>'
            % ("" if w in r["wheels"] else " off", WHEELS[w]["numeral"], WHEELS[w]["numeral"])
            for w in (1, 2, 3)) if r["wheels"] else ""
        return (
            "<tr class=prow data-uid=\"%d\" title=\"Click for this player's spins and assets\">"
            "<td data-v=\"%s\"><span class=who><span class=\"dot%s\" title=\"%s\"></span>@%s"
            "</span>%s%s</td>"
            "<td data-v=\"%d\" title=\"%s\">%s</td>"
            "<td class=mono data-v=\"%d\">%s</td>"
            "<td class=mono data-v=\"%d\">%s</td>"
            "<td class=mono data-v=\"%d\">%s</td>"
            "<td class=mono data-v=\"%d\">%s</td>"
            "<td data-v=\"%s\">%s</td>"
            "<td data-v=\"%d\">%s</td></tr>"
            % (r["id"],
               esc(str(r["display"]).lower()),
               " live" if r.get("active") else "",
               "on the page now" if r.get("active") else "not on the page",
               esc(r["display"]),
               (' <span class="pill no">\U0001F621 %d</span>' % r["forfeits"]) if r["forfeits"] else "",
               seat,
               r["rank_ord"], esc((r["rank"] or "") + (" \u00b7 " + r["title"] if r["title"] else "")),
               esc(r["rank"] or "-"),
               r["bank"], f"{r['bank']:,}",
               r["delta"], f"{r['delta']:+,}",
               r["available"], f"{r['available']:,}",
               r["spins"], r["spins"],
               esc(r.get("last_at") or ""), last,
               1 if r["in_bank"] else 0,
               '<span class="pill ok">in bank</span>' if r["in_bank"] else
               '<span class="pill no">not in bank</span>'))

    body = "".join(row_html(r) for r in rows)
    prize_body = "".join(
        "<tr><td data-v=\"%s\">@%s</td><td data-v=\"%s\">%s %s</td>"
        "<td class=mono data-v=\"%d\">%s</td></tr>"
        % (esc(str(p["display"]).lower()), esc(p["display"]),
           esc(p["name"].lower()), p["emoji"], esc(p["name"]), p["qty"], f"{p['qty']:,}")
        for p in prizes)
    if not prize_body:
        prize_body = ('<tr data-fixed=1><td colspan=3 class=hint style="padding:14px 6px">'
                      'Nothing owed right now.</td></tr>')
    player_head = "".join(
        '<th class=sortable data-col="%d" data-type="%s"><button type=button>%s'
        '<span class=arrow>\u21C5</span></button></th>' % (i, kind, esc(label))
        for i, (label, kind) in enumerate(PLAYER_COLS))
    prize_head = "".join(
        '<th class=sortable data-col="%d" data-type="%s"><button type=button>%s'
        '<span class=arrow>\u21C5</span></button></th>' % (i, kind, esc(label))
        for i, (label, kind) in enumerate(PRIZE_COLS))
    cleared = nicetime(meta.get("prizes_since") or "")
    cleared_note = ("Cleared " + cleared + "." if cleared else
                    "Nothing cleared yet, so this is everything ever won.")
    unlock_btn = ("Every wheel is open for you." if meta.get("me_unlocked")
                  else '<form method=post action="/admin/unlock" style="display:inline">'
                       '<button>Unlock all wheels for me</button></form>')
    return page("""
<header class=top><div class=brand>Real Humans \u00b7 <b>admin</b></div>
<div class=spacer></div><span class=pill>%s players in the bank</span>
<span class=pill>read %s</span><a class=btn href="/">Site</a></header>
<div class=card style="margin-bottom:18px">
  <div class=lab>Since your last update</div>
  <p class=hint>Everything you have to do is in one file: <b>Deltas + prizes</b>. Type each
  delta into column H of the sheet (that moves J, and J - P x 1e6 is the spending figure),
  hand out the prizes in the three logo columns and the eight asset columns after them, then
  <b>Sync bank</b> and <b>Mark applied</b>. Nothing is written for you. Mark applied sets both
  lists back to zero: the Deltas file and the prize obligations below. It never clears a
  fragment. Someone marked <i>not in bank</i> has no BITS, so they cannot spin yet.</p>
  <form method=post action="/admin/sync" style="display:inline"><button class=primary>Sync bank</button></form>
  <form method=post action="/admin/apply" style="display:inline"
    onsubmit="return confirm('Deltas typed into the sheet and prizes handed out? This clears both lists.')">
    <button>Mark applied</button></form>
  <a class=btn href="/admin/deltas.csv">Deltas + prizes</a>
  <a class=btn href="/admin/spins.csv">Every spin</a>
  <a class=btn href="/admin/prizes.csv">Prizes only</a>
  <form method=post action="/admin/upload" enctype=multipart/form-data style="margin-top:14px">
    <div class=lab>Or upload the workbook</div>
    <input type=file name=file accept=.xlsx style="margin:8px 0"><button>Upload and sync</button></form>
  <hr style="border:0;border-top:1px solid var(--line);margin:16px 0">
  <div class=lab>Jackpot and testing</div>
  <p class=hint style="margin:6px 0 10px">Wheel III pot: <b>%s BITS</b>. Every paid Wheel III
  spin adds 30,000 to it and the winner takes all of it. Resetting it hands nothing to anybody.
  Unlocking every wheel is for your own account, so you can try Wheel II and III without
  spinning a thousand times.</p>
  <form method=post action="/admin/jackpot" style="display:inline"
    onsubmit="return confirm('Set the Wheel III pot back to zero?')"><button>Reset jackpot</button></form>
  %s
</div>
<div class="card" style="margin-bottom:18px"><div class=lab>Prize obligations</div>
  <p class=hint>The three logo prizes and every booster and token won on any wheel, owed since
  the last <b>Mark applied</b>. Fragments are not here on purpose: they are the player's for
  good, and you can see them by clicking the player below. <i>%s</i></p>
  <table id=prizeTbl><thead><tr>%s</tr></thead><tbody>%s</tbody></table></div>
<div class=card><div class=lab>Players</div>
  <p class=hint>Click a heading to sort: the first click puts the biggest number on top, the
  second turns it round. Rank sorts by where the rank sits on the ladder, not by the alphabet.
  A green dot means that player has the page open right now; grey means they do not.
  <span id=liveNote class=hint></span></p>
  <table id=playersTbl><thead><tr>%s</tr></thead><tbody>%s</tbody></table></div>
<footer>Admin is the Lichess account @CC_Prestiged.</footer>
<script>%s</script>
""" % (meta["bank_players"], esc(meta["bank_synced_at"] or "never"), f"{meta['jackpot']:,}",
       unlock_btn, cleared_note, prize_head, prize_body, player_head, body, SORT_JS)
             + PLAYER_MODAL, wheel=0)        # the admin keeps the old gold accent

PLAYER_MODAL = """
<style>
.prow{cursor:pointer}
.prow:hover td{background:rgba(255,255,255,.04)}
.wtab{display:inline-block;min-width:18px;text-align:center;font-size:.66rem;
border:1px solid rgba(255,255,255,.22);border-radius:5px;padding:0 3px;margin-left:5px;
color:var(--muted);line-height:1.5}
.wtab.off{opacity:.32}
#pmodal{position:fixed;inset:0;z-index:60;display:none;padding:28px 16px;overflow:auto;
  background:rgba(6,8,13,.86);backdrop-filter:blur(4px)}
#pmodal.on{display:block}
.pm-card{max-width:1000px;margin:0 auto;background:#11151e;border:1px solid var(--line);
  border-radius:18px;padding:22px}
.pm-head{display:flex;align-items:center;gap:10px;flex-wrap:wrap;margin-bottom:4px}
.pm-name{font-family:Georgia,serif;font-size:1.35rem}
.pm-x{margin-left:auto;padding:6px 13px;border-radius:10px}
.pm-sum{display:grid;grid-template-columns:repeat(auto-fit,minmax(112px,1fr));gap:10px;margin:16px 0}
.pm-sum div{border:1px solid var(--line);border-radius:12px;padding:10px 12px}
.pm-sum span{display:block;font-size:.66rem;letter-spacing:.1em;text-transform:uppercase;
  color:var(--muted)}
.pm-sum b{font-size:1.05rem;font-variant-numeric:tabular-nums}
.pm-sec{margin:18px 0 6px}
.pm-two{display:grid;grid-template-columns:1fr 1fr;gap:18px}
@media(max-width:760px){.pm-two{grid-template-columns:1fr}}
.frag{display:grid;grid-template-columns:24px minmax(0,1fr) auto;gap:8px;align-items:center;
  padding:6px 0;border-bottom:1px solid rgba(255,255,255,.05);font-size:.88rem}
.pips{display:inline-flex;gap:4px;margin-left:8px;vertical-align:middle}
.pips i{width:9px;height:9px;border-radius:50%;background:rgba(255,255,255,.14);
  display:inline-block}
.pips i.on{background:var(--acc);box-shadow:0 0 7px var(--glow)}
.pm-table{max-height:46vh;overflow:auto;border-top:1px solid var(--line)}
.pm-table th{position:sticky;top:0;background:#11151e;z-index:1}
.pm-bad{color:#fca5a5}
.pm-wheels{display:flex;gap:8px;flex-wrap:wrap;margin-top:6px}
.pm-wheels div{border:1px solid var(--line);border-radius:12px;padding:8px 12px;font-size:.84rem}
.pm-wheels b{color:var(--acc)}
.pm-wheels .off{opacity:.55}
</style>
<div id=pmodal role=dialog aria-modal=true aria-label="Player">
  <div class=pm-card>
    <div class=pm-head>
      <span class=pm-name id=pmName>&mdash;</span>
      <span class=pill id=pmRank></span>
      <span class=pill id=pmTitle></span>
      <button class=pm-x id=pmClose type=button>Close &times;</button>
    </div>
    <div class=pm-wheels id=pmWheels></div>
    <div class=pm-sum id=pmSum></div>
    <div class=pm-two>
      <div>
        <div class="lab pm-sec">Fragments, kept for good</div>
        <div id=pmFrags></div>
      </div>
      <div>
        <div class="lab pm-sec">Owed now</div>
        <div id=pmOwed></div>
        <div class="lab pm-sec">Won in total</div>
        <div id=pmEver class=hint></div>
      </div>
    </div>
    <div class="lab pm-sec">Spins</div>
    <div class=pm-table>
      <table>
        <thead><tr><th>When</th><th>Wheel</th><th>Wager</th><th>Landed</th><th>Status</th>
          <th>Cost</th><th>Paid / owed</th></tr></thead>
        <tbody id=pmBody></tbody>
      </table>
    </div>
  </div>
</div>
<script>
(function(){
  var pm = document.getElementById('pmodal');
  function esc(v){
    return String(v == null ? '' : v).replace(/[&<>"]/g, function(c){
      return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];
    });
  }
  function money(n){ return Number(n || 0).toLocaleString(); }
  function badge(w){
    return '<span class="wtag w' + w + '">' + ['', 'I', 'II', 'III'][w || 1] + '</span>';
  }
  function list(box, pairs, empty){
    if (!pairs.length) { box.innerHTML = '<div class=hint>' + empty + '</div>'; return; }
    box.innerHTML = pairs.map(function(p){
      return '<div class=frag><span>' + p[0] + '</span><b>' + esc(p[1]) + '</b>' +
        '<span class=mono>' + money(p[2]) + '</span></div>';
    }).join('');
  }
  async function open(uid){
    var r = await fetch('/admin/player/' + uid + '.json');
    if (!r.ok) { alert('Could not load that player.'); return; }
    var d = await r.json(), s = d.summary;
    document.getElementById('pmName').textContent = '@' + d.display;
    document.getElementById('pmRank').textContent = d.rank || 'no rank';
    document.getElementById('pmTitle').textContent = d.title || 'no title';
    document.getElementById('pmWheels').innerHTML = d.progress.map(function(p){
      var open = p.unlocked;
      return '<div class="' + (open ? '' : 'off') + '">' + p.numeral + ' \u00b7 ' + p.name +
        '<br><b>' + (open ? 'open' : 'locked') + '</b> \u00b7 ' +
        money(p.have) + ' / ' + money(p.need) + ' ' + p.unit + '</div>';
    }).join('');
    document.getElementById('pmSum').innerHTML = [
      ['Spins', money(s.spins)], ['Collected', money(s.collected)],
      ['Left early', money(s.forfeited)], ['Staked', money(s.staked)],
      ['Paid back', money(s.won)], ['Net', money(s.won - s.staked)],
      ['Biggest win', money(s.best)]
    ].map(function(p){
      return '<div><span>' + p[0] + '</span><b>' + p[1] + '</b></div>';
    }).join('');

    document.getElementById('pmFrags').innerHTML = d.fragments.length
      ? d.fragments.map(function(f){
          var pips = '';
          for (var i = 0; i < 3; i++) pips += '<i class="' + (i < f.count ? 'on' : '') + '"></i>';
          return '<div class=frag><span>' + f.emoji + '</span><b>' + esc(f.name) +
            '</b><span class=mono>' + f.count + ' / 3<span class=pips>' + pips + '</span></span></div>';
        }).join('')
      : '<div class=hint>No fragments yet.</div>';

    list(document.getElementById('pmOwed'),
         d.owed.map(function(o){ return [o.emoji, o.name, o.qty]; }),
         'Nothing owed right now.');
    document.getElementById('pmEver').innerHTML = d.ever.length
      ? d.ever.map(function(o){ return o.emoji + ' ' + esc(o.name) + ' \u00d7 ' + money(o.qty); })
          .join(' \u00b7 ')
      : 'Nothing ever won.';

    document.getElementById('pmBody').innerHTML = d.spins.map(function(x){
      var wtag = badge(x.wheel);
      var what;
      if (x.status === 'settled') {
        what = x.emoji + ' ' + esc(x.label);
        if (x.kind === 'item') what += ' &times; ' + money(x.prize_qty) + (x.note ? ' (' + esc(x.note) + ')' : '');
        if (x.kind === 'free') what += ' &rarr; ' + money(x.prize_qty) + ' free';
      } else if (x.status === 'forfeited') {
        what = '\U0001F621 left';
      } else {
        what = 'still turning';
      }
      var paid = x.status === 'settled'
        ? (x.kind === 'item' || x.kind === 'free' ? money(x.prize_qty) + ' prize'
           : '+' + money(x.payout))
        : '&mdash;';
      return '<tr><td class=mono>' + new Date(x.when).toLocaleString() + '</td>' +
        '<td>' + wtag + '</td><td class=mono>x' + money(x.wager) + (x.free ? ' free' : '') + '</td>' +
        '<td>' + what + '</td>' +
        '<td' + (x.status === 'forfeited' ? ' class=pm-bad' : '') + '>' + esc(x.status) + '</td>' +
        '<td class=mono>' + (x.cost ? '-' + money(x.cost) : 'free') + '</td>' +
        '<td class=mono>' + paid + '</td></tr>';
    }).join('') || '<tr><td colspan=7 class=hint>No spins yet.</td></tr>';
    pm.classList.add('on');
  }
  document.getElementById('pmClose').addEventListener('click', function(){ pm.classList.remove('on'); });
  pm.addEventListener('click', function(e){ if (e.target === pm) pm.classList.remove('on'); });
  document.addEventListener('keydown', function(e){ if (e.key === 'Escape') pm.classList.remove('on'); });
  document.querySelectorAll('tr.prow').forEach(function(tr){
    tr.addEventListener('click', function(){ open(tr.dataset.uid); });
  });
  makeSortable(document.getElementById('playersTbl'), 'bs_admin_sort');
  makeSortable(document.getElementById('prizeTbl'), 'bs_prize_sort');

  /* ---- who is on the page: a green dot that pulses, updated in place so the order holds */
  (function(){
    var note = document.getElementById('liveNote');
    function tick(){
      fetch('/admin/live.json', {cache: 'no-store'}).then(function(r){
        return r.ok ? r.json() : null;
      }).then(function(d){
        if (!d) return;
        var count = 0;
        document.querySelectorAll('tr.prow').forEach(function(tr){
          var info = d.players[tr.getAttribute('data-uid')] || {};
          var dot = tr.querySelector('.dot');
          if (!dot) return;
          var live = !!info.active;
          if (live) count++;
          dot.classList.toggle('live', live);
          dot.title = live ? 'on the page now' :
            (info.seen_ago ? ('last seen ' + info.seen_ago) : 'has not opened the page yet');
        });
        if (note) note.textContent = count
          ? ('\u00b7 ' + count + ' on the page right now \u00b7 updated ' + new Date().toLocaleTimeString())
          : ('\u00b7 nobody has the page open \u00b7 updated ' + new Date().toLocaleTimeString());
      }).catch(function(){});
    }
    setInterval(tick, 15000);
    tick();
  })();
})();
</script>
"""

# --------------------------------------------------------------------------- app
app = FastAPI(docs_url=None, redoc_url=None, title="Real Humans · Mystery BITS")


def current_user(request, db):
    blob = unsign(request.cookies.get(COOKIE))
    if not blob:
        return None
    try:
        uid = int(blob.get("uid") or 0)
    except (TypeError, ValueError):
        return None
    return db.get(User, uid) if uid else None


def is_admin(user):
    if not user:
        return False
    if user.is_admin:                       # only ever set by a bank sync or by you, in the DB
        return True
    return not user.password_hash and name_key(user.lichess_id) == ADMIN_LICHESS


def login_cookie(resp, user):
    resp.set_cookie(COOKIE, sign({"uid": user.id}), httponly=True, samesite="lax",
                    max_age=60 * 60 * 24 * 30, secure=BASE_URL.startswith("https"))
    return resp


def ensure_admin_flag(db, user):
    """Called on every successful login. The admin is never left waiting for approval:
    before the first bank upload there is nobody else to approve them."""
    if user.password_hash or name_key(user.lichess_id) != ADMIN_LICHESS:
        return
    changed = False
    if not user.is_admin:
        user.is_admin, changed = True, True
    if not user.approved:
        user.approved, changed = True, True
    if changed:
        db.commit()


@app.on_event("startup")
def _startup():
    Base.metadata.create_all(engine)
    ensure_columns()
    db = SessionLocal()
    try:
        backfill_awards(db)
    finally:
        db.close()


def player_payload(db, user):
    """Everything the player page needs, in one object. The page switches wheels from this
    and never has to fetch a wheel it has not seen yet."""
    counts = wheel_counts(db, user.id)
    have = unlocked_wheels(db, user, counts)
    state = state_of(db, user)
    wheels = []
    for w in (1, 2, 3):
        spec = WHEELS[w]
        avg = average_return(w)
        wheels.append({
            "id": w, "numeral": spec["numeral"], "name": spec["name"], "tag": spec["tag"],
            "mood": spec["mood"], "cost": spec["cost"], "maxWager": wheel_max_wager(w),
            "wager": bool(spec["wager"]), "unlocked": w in have,
            "progress": progress_for(db, user, w, counts),
            "average": {"per": round(avg["per_spin"], 2), "cost": avg["cost"],
                        "pct": round(avg["pct"], 1)},
            "slices": [{"label": s["label"], "lines": s["lines"], "emoji": s["emoji"],
                        "code": s["code"], "kind": s["kind"], "amount": int(s["amount"]),
                        "color": s["color"], "ink": s["ink"], "note": s["note"],
                        "pct": round(int(s["weight"]) / 100.0, 2)} for s in spec["slices"]],
            "free": free_state(db, user, w),
        })
    return {"wheels": wheels, "active": 1, "multi": state["multiple"],
            "available": state["available"], "bank": state["bank"], "delta": state["delta"],
            "display": user.display or user.lichess_id, "jackpot": jackpot_value(db),
            "ping": PING_EVERY, "reveal": REVEAL_AFTER, "ttl": SPIN_TTL,
            "is_admin": is_admin(user)}


@app.get("/emoji-sheet.png")
def emoji_sheet():
    blob = base64.b64decode(EMOJI_SHEET.split(",", 1)[1])
    return Response(blob, media_type="image/png",
                    headers={"Cache-Control": "public, max-age=31536000"})


@app.get("/healthz")
def healthz():
    return {"ok": True, "wheels": len(WHEELS)}


@app.get("/", response_class=Response)
def home(request: Request, db=Depends(get_db)):
    sweep(db)
    user = current_user(request, db)
    if not user:
        return Response(login_page(), media_type="text/html")
    touch(db, user, force=True)
    payload = player_payload(db, user)
    return Response(index_page(user, payload, payload["delta"]), media_type="text/html")


# --- auth
@app.get("/auth/lichess")
def auth_lichess():
    verifier, challenge = pkce()
    state = secrets.token_urlsafe(16)
    resp = RedirectResponse(lichess_url(state, challenge))
    resp.set_cookie("bs_oauth", sign({"state": state, "verifier": verifier}),
                    httponly=True, samesite="lax", max_age=600)
    return resp


@app.get("/auth/callback")
def auth_callback(request: Request, code: str = "", state: str = "", db=Depends(get_db)):
    blob = unsign(request.cookies.get("bs_oauth"))
    if not blob or blob.get("state") != state:
        return Response(login_page("That login link expired. Try again."), media_type="text/html",
                        status_code=400)
    try:
        acc = lichess_account(lichess_token(code, blob["verifier"])["access_token"])
    except LichessError as exc:
        return Response(login_page("Lichess did not finish the login: %s." % exc),
                        media_type="text/html", status_code=502)
    except Exception as exc:                              # noqa: BLE001
        return Response(login_page("Lichess did not finish the login (%s)." % exc),
                        media_type="text/html", status_code=502)
    lid = (acc.get("id") or "").lower()
    user = db.execute(select(User).where(User.lichess_id == lid)).scalar_one_or_none()
    if not user:
        user = User(lichess_id=lid, display=acc.get("username") or lid)
        db.add(user)
        db.commit()
    apply_bank_to_user(db, user)
    ensure_admin_flag(db, user)
    return login_cookie(RedirectResponse("/", status_code=303), user)


@app.post("/auth/local")
def auth_local(username: str = Form(...), password: str = Form(...), mode: str = Form("login"),
               db=Depends(get_db)):
    if not ALLOW_LOCAL:
        return Response(login_page("Use the Lichess button. Local accounts are turned off."),
                        media_type="text/html", status_code=403)
    lid = username.strip().lstrip("@").lower()
    if name_key(lid) == ADMIN_LICHESS:
        return Response(login_page("The admin account must be signed in with Lichess."),
                        media_type="text/html", status_code=403)
    if not lid or " " in lid:
        return Response(login_page("Use your Lichess username, without spaces."),
                        media_type="text/html", status_code=400)
    user = db.execute(select(User).where(User.lichess_id == lid)).scalar_one_or_none()
    if mode == "register":
        if user and user.password_hash:
            return Response(login_page("That name already has an account. Log in instead."),
                            media_type="text/html", status_code=400)
        if not user:
            user = User(lichess_id=lid, display=username.strip().lstrip("@"))
            db.add(user)
        user.password_hash = hash_password(password)
        db.commit()
        apply_bank_to_user(db, user)
        ensure_admin_flag(db, user)
    else:
        if not user or not user.password_hash or not check_password(password, user.password_hash):
            return Response(login_page("Wrong name or password."), media_type="text/html",
                            status_code=401)
    return login_cookie(RedirectResponse("/", status_code=303), user)


@app.get("/logout")
def logout():
    resp = RedirectResponse("/", status_code=303)
    resp.delete_cookie(COOKIE)
    return resp


# --- player api
@app.get("/api/me")
def api_me(request: Request, db=Depends(get_db)):
    user = current_user(request, db)
    if not user:
        return JSONResponse({"detail": "not signed in"}, status_code=401)
    touch(db, user)
    return player_payload(db, user)


@app.get("/api/ping")
def api_ping(request: Request, db=Depends(get_db)):
    """The open page says hello, so the admin can see who is actually there."""
    user = current_user(request, db)
    if not user:
        return JSONResponse({"ok": False}, status_code=401)
    touch(db, user, force=True)
    return {"ok": True, "now": utcnow().isoformat(), "active_seconds": ACTIVE_SECONDS}


@app.post("/api/spin")
async def api_spin(request: Request, db=Depends(get_db)):
    user = current_user(request, db)
    if not user:
        return JSONResponse({"detail": "not signed in"}, status_code=401)
    touch(db, user)
    try:
        body = await request.json()
    except Exception:                                     # noqa: BLE001
        body = {}
    if not isinstance(body, dict):
        body = {}
    use_free = bool(body.get("free"))
    try:
        spin, cost, balance = start_spin(db, user, body.get("wager"), free=use_free,
                                         wheel=body.get("wheel") or 1)
    except Refused as exc:
        return JSONResponse({"detail": exc.message}, status_code=400)
    return {"spin_id": spin.id, "cost": cost, "balance": balance, "commit": spin.commit,
            "wager": spin.wager, "wheel": int(spin.wheel or 1), "capped": spin.capped,
            "ttl": SPIN_TTL, "free": use_free,
            "free_spins": free_state_all(db, user),
            "reveal_after": REVEAL_AFTER, "expires_at": spin.expires_at.isoformat()}


@app.post("/api/spin/{spin_id}/reveal")
def api_reveal(spin_id: str, request: Request, db=Depends(get_db)):
    """Where it stopped. Nothing is paid yet."""
    user = current_user(request, db)
    if not user:
        return JSONResponse({"detail": "not signed in"}, status_code=401)
    try:
        spin = reveal_spin(db, user, spin_id)
    except Refused as exc:
        return JSONResponse({"detail": exc.message}, status_code=400)
    payload = {"index": spin.index_landed, "label": spin.label, "kind": spin.kind,
               "payout": spin.payout, "prize_qty": spin.prize_qty, "wager": spin.wager,
               "wheel": int(spin.wheel or 1), "seed": spin.seed, "commit": spin.commit,
               "check": {"sha256_seed": commit_of(spin.seed) == spin.commit,
                         "index_from_seed": pick_index(int(spin.wheel or 1), spin.seed)}}
    if spin.kind == "jackpot":
        payload["pot"] = jackpot_value(db)
    return payload


@app.post("/api/spin/{spin_id}/claim")
def api_claim(spin_id: str, request: Request, db=Depends(get_db)):
    """Collect it. If the page was closed before this call, the spin is a forfeit."""
    user = current_user(request, db)
    if not user:
        return JSONResponse({"detail": "not signed in"}, status_code=401)
    try:
        spin, fresh = claim_spin(db, user, spin_id)
    except Refused as exc:
        return JSONResponse({"detail": exc.message}, status_code=400)
    return {"paid": spin.payout, "prize_qty": spin.prize_qty, "label": spin.label,
            "kind": spin.kind, "index": spin.index_landed, "wager": spin.wager,
            "wheel": int(spin.wheel or 1), "note": spin.note or "",
            "balance": state_of(db, user)["available"], "replayed": not fresh,
            "free": spin.cost == 0, "free_spins": free_state_all(db, user),
            "seed": spin.seed, "commit": spin.commit, "status": spin.status,
            "jackpot": jackpot_value(db)}


@app.get("/api/spins")
def api_spins(request: Request, db=Depends(get_db)):
    user = current_user(request, db)
    if not user:
        return JSONResponse({"detail": "not signed in"}, status_code=401)
    sweep(db)
    rows = db.execute(select(Spin).where(Spin.user_id == user.id)
                      .order_by(Spin.created_at.desc()).limit(200)).scalars().all()
    return [{"id": s.id, "created_at": s.created_at.isoformat(), "wager": s.wager,
             "wheel": int(s.wheel or 1), "badge": wheel_badge(s.wheel),
             "cost": s.cost, "status": s.status, "label": s.label or "",
             "kind": s.kind or "", "payout": s.payout or 0, "index": s.index_landed,
             "prize_qty": s.prize_qty or 0, "note": s.note or "",
             "free": s.cost == 0} for s in rows]


# --- admin
@app.get("/admin", response_class=Response)
def admin_home(request: Request, db=Depends(get_db)):
    who = admin_guard(request, db)
    if not who:
        user = current_user(request, db)
        if user:
            message = ("You are signed in as @%s, but this page belongs to @%s. If that is the "
                       "wrong name, fix BS_ADMIN_LICHESS in your host's environment settings."
                       % (user.lichess_id, ADMIN_RAW.strip().lstrip("@")))
        else:
            message = "Admin is the Lichess account @%s. Log in with it." % ADMIN_RAW.strip().lstrip("@")
        return Response(login_page(message), media_type="text/html", status_code=403)
    since = prizes_since(db)
    meta = {"bank_synced_at": meta_get(db, "bank_synced_at", ""),
            "bank_players": len(bank_rows_list(db)),
            "prizes_since": since.isoformat() if since else "",
            "jackpot": jackpot_value(db),
            "me_unlocked": bool(who and who.test_unlocked)}
    return Response(admin_page(admin_rows(db), prize_rows(db), meta), media_type="text/html")


@app.post("/admin/sync")
def admin_sync(request: Request, db=Depends(get_db)):
    if not admin_guard(request, db):
        return Response("forbidden", status_code=403)
    try:
        sync_bank(db)
    except BankError as exc:
        return Response('<p style="font-family:sans-serif">%s</p><p><a href="/admin">back</a></p>'
                        % esc(exc), media_type="text/html", status_code=400)
    return RedirectResponse("/admin", status_code=303)


@app.post("/admin/upload")
async def admin_upload(request: Request, file: UploadFile = File(...), db=Depends(get_db)):
    if not admin_guard(request, db):
        return Response("forbidden", status_code=403)
    blob = await file.read()
    try:
        parse_bank(blob)                                   # fail loudly before storing
    except BankError as exc:
        return Response(esc(str(exc)), media_type="text/html", status_code=400)
    sync_bank(db, blob=blob, origin="upload")
    return RedirectResponse("/admin", status_code=303)


@app.post("/admin/apply")
def admin_apply(request: Request, db=Depends(get_db)):
    if not admin_guard(request, db):
        return Response("forbidden", status_code=403)
    now = utcnow()
    for u in db.query(User).all():
        last = db.execute(select(func.max(Ledger.id)).where(Ledger.user_id == u.id)).scalar()
        u.applied_through = int(last or 0)
    meta_set(db, "prizes_applied_at", now.isoformat())      # prize list starts empty again
    db.commit()
    return RedirectResponse("/admin", status_code=303)


@app.post("/admin/unlock")
def admin_unlock(request: Request, db=Depends(get_db)):
    """Testing switch: give the admin account every wheel at once."""
    who = admin_guard(request, db)
    if not who:
        return Response("forbidden", status_code=403)
    who.test_unlocked = not bool(who.test_unlocked)
    db.commit()
    return RedirectResponse("/admin", status_code=303)


@app.post("/admin/jackpot")
def admin_jackpot(request: Request, db=Depends(get_db)):
    if not admin_guard(request, db):
        return Response("forbidden", status_code=403)
    meta_set(db, "jackpot", "0")
    db.commit()
    return RedirectResponse("/admin", status_code=303)


@app.get("/admin/live.json")
def admin_live(request: Request, db=Depends(get_db)):
    """Who has the page open right now, for the dots on the admin table."""
    if not admin_guard(request, db):
        return JSONResponse({"detail": "forbidden"}, status_code=403)
    now = utcnow()
    out = {}
    for u in db.query(User).all():
        out[str(u.id)] = {"active": is_active(u, now), "seen": seen_iso(u),
                          "seen_ago": ago(seen_iso(u)) if u.last_seen else ""}
    return {"now": now.isoformat(), "active_seconds": ACTIVE_SECONDS, "players": out}


@app.get("/admin/player/{uid}.json")
def admin_player_json(uid: int, request: Request, db=Depends(get_db)):
    if not admin_guard(request, db):
        return JSONResponse({"detail": "forbidden"}, status_code=403)
    user = db.get(User, uid)
    if not user:
        return JSONResponse({"detail": "no such player"}, status_code=404)
    spins = db.execute(select(Spin).where(Spin.user_id == uid)
                       .order_by(Spin.created_at.desc()).limit(600)).scalars().all()
    total = db.execute(select(func.count(Spin.id)).where(Spin.user_id == uid)).scalar()
    collected, staked, won = db.execute(
        select(func.count(Spin.id), func.coalesce(func.sum(Spin.cost), 0),
               func.coalesce(func.sum(Spin.payout), 0))
        .where(Spin.user_id == uid, Spin.status == "settled")).one()
    forfeited = db.execute(select(func.count(Spin.id))
                           .where(Spin.user_id == uid, Spin.status == "forfeited")).scalar()
    best = db.execute(select(func.coalesce(func.max(Spin.payout), 0))
                      .where(Spin.user_id == uid, Spin.status == "settled")).scalar()

    counts = wheel_counts(db, uid)
    have = unlocked_wheels(db, user, counts)
    progress = []
    for w in (1, 2, 3):
        p = progress_for(db, user, w, counts)
        progress.append({"id": w, "numeral": WHEELS[w]["numeral"], "name": WHEELS[w]["name"],
                         "unlocked": w in have, "have": p["have"], "need": p["need"],
                         "unit": p["unit"]})

    frags = fragments_of(db, uid)
    since = prizes_since(db)
    owed = award_totals(db, user_id=uid, since=since).get(uid, {})
    ever = award_totals(db, user_id=uid).get(uid, {})
    row = bank_row(db, user.lichess_id) or {}

    def pack(mapping):
        return [{"key": k, "name": item_name(k), "emoji": item_emoji(k), "qty": int(q)}
                for k, q in sorted(mapping.items(),
                                   key=lambda kv: PRIZE_KEYS.index(kv[0])
                                   if kv[0] in PRIZE_KEYS else 99) if q]

    return {
        "display": user.display or user.lichess_id,
        "lichess_id": user.lichess_id,
        "rank": user.bank_rank or "",
        "title": row.get("title", "") or "",
        "in_bank": bank_row(db, user.lichess_id) is not None,
        "bank": int(user.bank_bits or 0),
        "delta": delta_for(db, uid, user.applied_through or 0),
        "free": free_state_all(db, user),
        "test_unlocked": bool(user.test_unlocked),
        "progress": progress,
        "counts": {str(k): v for k, v in counts.items()},
        "fragments": [{"key": k, "name": "%s fragment" % item_name(FRAGMENTS[k]),
                       "emoji": item_emoji(FRAGMENTS[k]), "count": int(frags.get(k, 0))}
                      for k in FRAGMENT_ORDER if frags.get(k, 0)],
        "owed": pack(owed),
        "ever": pack(ever),
        "summary": {"spins": int(total or 0), "collected": int(collected or 0),
                    "staked": int(staked or 0), "won": int(won or 0),
                    "forfeited": int(forfeited or 0), "best": int(best or 0)},
        "spins": [{"when": x.created_at.isoformat(), "wager": x.wager, "cost": x.cost,
                   "wheel": int(x.wheel or 1), "label": x.label or "",
                   "emoji": emoji_of(x.label), "kind": x.kind or "",
                   "payout": x.payout or 0, "prize_qty": x.prize_qty or 0,
                   "note": x.note or "", "free": x.cost == 0, "status": x.status}
                  for x in spins],
    }


@app.get("/admin/deltas.csv")
def admin_deltas(request: Request, db=Depends(get_db)):
    """One line per player: the delta to type into their sheet, and everything owed.

    A to G are the figures the sheet needs. H to R are the items: three logos, then the
    seven boosters, then Legacy Tokens. Fragments are not here, on purpose: they are the
    player's own and Mark applied must not touch them."""
    if not admin_guard(request, db):
        return Response("forbidden", status_code=403)
    since = prizes_since(db)
    totals = award_totals(db, since=since)
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["username", "rank", "spending_before", "delta_to_type_into_H",
                "spending_after", "spins", "staked"] +
               [ITEMS[k]["column"] for k in PRIZE_KEYS])
    for r in admin_rows(db):
        mine = totals.get(r["id"], {})
        counts = [int(mine.get(k, 0)) for k in PRIZE_KEYS]
        if r["delta"] == 0 and r["spins"] == 0 and not any(counts):
            continue
        w.writerow(["@" + str(r["display"]).lstrip("@"), r["rank"], r["bank"], r["delta"],
                    r["available"], r["spins"], r["staked"]] + counts)
    return PlainTextResponse(buf.getvalue(), media_type="text/csv",
                             headers={"Content-Disposition": 'attachment; filename="rh-deltas.csv"'})


@app.get("/admin/spins.csv")
def admin_spins(request: Request, db=Depends(get_db)):
    if not admin_guard(request, db):
        return Response("forbidden", status_code=403)
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["when", "username", "wheel", "wager", "capped", "cost", "free_spin", "status",
                "landed", "kind", "bits_paid", "prize_owed", "note", "seed", "commit"])
    q = select(Spin, User).join(User, Spin.user_id == User.id).order_by(Spin.created_at)
    for spin, user in db.execute(q).all():
        w.writerow([spin.created_at.isoformat(), "@" + (user.lichess_id or ""),
                    WHEELS.get(int(spin.wheel or 1), WHEELS[1])["numeral"], spin.wager,
                    "yes" if spin.capped else "", spin.cost,
                    "yes" if spin.cost == 0 else "", spin.status, spin.label or "",
                    spin.kind or "", spin.payout or 0, spin.prize_qty or 0, spin.note or "",
                    spin.seed or "", spin.commit or ""])
    return PlainTextResponse(buf.getvalue(), media_type="text/csv",
                             headers={"Content-Disposition": 'attachment; filename="rh-spins.csv"'})


@app.get("/admin/prizes.csv")
def admin_prizes(request: Request, db=Depends(get_db)):
    if not admin_guard(request, db):
        return Response("forbidden", status_code=403)
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["username", "prize", "owed"])
    for p in prize_rows(db):
        w.writerow(["@" + str(p["display"]).lstrip("@"), p["name"], p["qty"]])
    return PlainTextResponse(buf.getvalue(), media_type="text/csv",
                             headers={"Content-Disposition": 'attachment; filename="rh-prizes.csv"'})


# --------------------------------------------------------------------------- emoji fallback
# Tiles in this order, one per emoji the three wheels use, 72 px each. Twemoji artwork,
# CC-BY 4.0. Only used when a browser has no colour emoji font of its own.
EMOJI_CODES = ["1f7e2", "1f331", "1f4a7", "1f539", "26a1", "1f49c", "1f3af", "1f36d",
               "1f4bc", "1f4b0", "1f3c6", "1f451", "1f5bc", "1f39e", "1f9ec", "1f504",
               "1f480", "1fa99", "1f4b5", "1f48e", "1f4a0", "1f525", "1f680", "1f31f",
               "1f409", "1f30c", "1f3b0", "1f45b", "1fa8e", "2604", "1f6e1", "1f52e"]

EMOJI_SHEET = "data:image/png;base64,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"


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))
