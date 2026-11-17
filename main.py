"""
Real Humans · Mystery BITS
==========================

One file. FastAPI + SQLAlchemy + openpyxl. Nothing else, because it has to be
pasted into a web editor by hand.

Rules this file implements
--------------------------
* spending BITS      = J - P * 1,000,000     (per row of the BITS Bank sheet)
* a spin costs       = 300 * wager
* wager in multiples = Sovereign and above only; below that the wager is forced
                       to 1 and they simply pay 300 for whatever the wheel gives
* prizes multiply    = a wager of 30,000 landing on "Moving Logo" owes 30,000
                       moving logos, and the admin export says so. The one
                       exception is "2 More Spins", which is always exactly two
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
                        Text, create_engine, func, select)
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
SAC_UNIT = 1_000_000                          # P is "BITS sacrificed, in millions"

SPIN_COST = int(env("BS_SPIN_COST", "300"))
SPIN_TTL = int(env("BS_SPIN_TTL_SECONDS", "30"))
# The wheel must actually turn before the result is revealed. Anything closed before this
# many seconds is a forfeit, which is the whole point of the timer.
REVEAL_AFTER = float(env("BS_REVEAL_SECONDS", "3.6"))
FREE_SPIN_AWARD = 2                    # "2 More Spins" hands over exactly two
MAX_WAGER = int(env("BS_MAX_WAGER", "1000000000"))
WAGER_MIN_RANK = env("BS_WAGER_MIN_RANK", "Sovereign")

# The real wheel, from wheelofnames.com/2px-y84: same 16 slices, same order,
# same colours, same emoji, same weights.
WHEEL = [
    {"label": "50 BITS", "emoji": "\U0001F7E2", "code": "1f7e2", "amount": 50,
     "kind": "bits", "weight": 1, "color": "#a8e6cf", "scales": True},
    {"label": "75 BITS", "emoji": "\U0001F331", "code": "1f331", "amount": 75,
     "kind": "bits", "weight": 1, "color": "#b2dfdb", "scales": True},
    {"label": "100 BITS", "emoji": "\U0001F4A7", "code": "1f4a7", "amount": 100,
     "kind": "bits", "weight": 1, "color": "#81d4fa", "scales": True},
    {"label": "150 BITS", "emoji": "\U0001F539", "code": "1f539", "amount": 150,
     "kind": "bits", "weight": 1, "color": "#4fc3f7", "scales": True},
    {"label": "200 BITS", "emoji": "\u26A1", "code": "26a1", "amount": 200,
     "kind": "bits", "weight": 1, "color": "#64b5f6", "scales": True},
    {"label": "250 BITS", "emoji": "\U0001F49C", "code": "1f49c", "amount": 250,
     "kind": "bits", "weight": 1, "color": "#ce93d8", "scales": True},
    {"label": "300 BITS", "emoji": "\U0001F3AF", "code": "1f3af", "amount": 300,
     "kind": "bits", "weight": 1, "color": "#ba68c8", "scales": True},
    {"label": "350 BITS", "emoji": "\U0001F36D", "code": "1f36d", "amount": 350,
     "kind": "bits", "weight": 1, "color": "#f06292", "scales": True},
    {"label": "400 BITS", "emoji": "\U0001F4BC", "code": "1f4bc", "amount": 400,
     "kind": "bits", "weight": 1, "color": "#f48fb1", "scales": True},
    {"label": "450 BITS", "emoji": "\U0001F4B0", "code": "1f4b0", "amount": 450,
     "kind": "bits", "weight": 1, "color": "#ffb74d", "scales": True},
    {"label": "500 BITS", "emoji": "\U0001F3C6", "code": "1f3c6", "amount": 500,
     "kind": "bits", "weight": 1, "color": "#ffd54f", "scales": True},
    {"label": "1000 BITS", "emoji": "\U0001F451", "code": "1f451", "amount": 1000,
     "kind": "bits", "weight": 1, "color": "#fff176", "scales": True},
    {"label": "Static Logo", "emoji": "\U0001F5BC\uFE0F", "code": "1f5bc", "amount": 1,
     "kind": "prize", "weight": 1, "color": "#ce93d8", "scales": True},
    {"label": "Moving Logo", "emoji": "\U0001F39E\uFE0F", "code": "1f39e", "amount": 1,
     "kind": "prize", "weight": 1, "color": "#ffb74d", "scales": True},
    {"label": "Custom Prefix", "emoji": "\U0001F9EC", "code": "1f9ec", "amount": 1,
     "kind": "prize", "weight": 1, "color": "#ba68c8", "scales": True},
    # two free spins, credited at the wager that won them; never multiplied
    {"label": "2 More Spins", "emoji": "\U0001F504", "code": "1f504", "amount": 2,
     "kind": "free", "weight": 1, "color": "#90caf9", "scales": False},
]

LEGACY_LADDER = ["Organiser", "Herald", "Sentinel", "Enforcer", "Sovereign", "Configurator",
                 "Administrator", "Bronze", "Silver", "Gold", "Platinum", "Osmium",
                 "Titanium", "Obsidian", "Diamond", "Founder"]


def norm(text) -> str:
    return "".join(ch for ch in str(text or "").lower() if "a" <= ch <= "z")


def rank_index(name) -> int:
    """The sheet is hand typed: match exactly, else the longest ladder name inside."""
    key = norm(name)
    if not key:
        return -1
    exact = {norm(step): i for i, step in enumerate(LEGACY_LADDER)}
    if key in exact:
        return exact[key]
    hits = [(len(norm(s)), i) for i, s in enumerate(LEGACY_LADDER) if norm(s) in key]
    return max(hits)[1] if hits else -1


def rank_of(name) -> str:
    i = rank_index(name)
    return LEGACY_LADDER[i] if i >= 0 else ""


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
    wager = Column(Integer)
    capped = Column(Boolean, default=False)          # wager forced down to 1
    cost = Column(BigInteger)
    status = Column(String(12), default="pending")   # pending | revealed | settled | forfeited
    commit = Column(String(64))
    seed = Column(String(64))
    index_landed = Column(Integer)
    label = Column(String(32))
    kind = Column(String(8))                         # bits | prize
    payout = Column(BigInteger, default=0)           # BITS paid
    prize_qty = Column(BigInteger, default=0)        # how many logos / spins, etc.
    created_at = Column(DateTime(timezone=True), default=utcnow)
    expires_at = Column(DateTime(timezone=True))
    settled_at = Column(DateTime(timezone=True))


class FreeSpin(Base):
    """Free spins are a credit, not a prize to hand out: `remaining` spins at `wager`,
    queued in the order they were won, and they stay put until they are used."""
    __tablename__ = "free_spins"
    id = Column(Integer, primary_key=True)
    user_id = Column(Integer, ForeignKey("users.id"), index=True)
    wager = Column(Integer)
    remaining = Column(Integer, default=0)
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


def free_grants(db, user_id):
    """Oldest grant first."""
    return db.execute(select(FreeSpin)
                      .where(FreeSpin.user_id == user_id, FreeSpin.remaining > 0)
                      .order_by(FreeSpin.created_at, FreeSpin.id)).scalars().all()


def free_state(db, user):
    grants = free_grants(db, user.id)
    if not grants:
        return None
    head = grants[0]
    return {"count": head.remaining, "wager": head.wager,
            "total": sum(g.remaining for g in grants),
            "queue": [{"wager": g.wager, "count": g.remaining} for g in grants]}


def grant_free_spins(db, user_id, wager, count=2):
    row = db.execute(select(FreeSpin)
                     .where(FreeSpin.user_id == user_id, FreeSpin.wager == wager)
                     .order_by(FreeSpin.created_at)).scalars().first()
    if row:
        row.remaining += count
    else:
        db.add(FreeSpin(user_id=user_id, wager=wager, remaining=count))


def meta_get(db, key, default=""):
    row = db.get(Meta, key)
    return row.v if row else default


def meta_set(db, key, value):
    db.merge(Meta(k=key, v=value))


# --------------------------------------------------------------------------- the wheel
def total_weight() -> int:
    return int(round(sum(slice_["weight"] for slice_ in WHEEL) * 1000))


def commit_of(seed: str) -> str:
    return hashlib.sha256(seed.encode("ascii")).hexdigest()


def pick_index(seed: str) -> int:
    n = int(seed, 16) % total_weight()
    run = 0
    for i, slice_ in enumerate(WHEEL):
        run += int(round(slice_["weight"] * 1000))
        if n < run:
            return i
    return len(WHEEL) - 1


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


def start_spin(db, user, wager, free=False):
    sweep(db)
    capped = False
    if free:
        grants = free_grants(db, user.id)
        if not grants:
            raise Refused("You have no free spins left.")
        grant = grants[0]
        wager, cost = grant.wager, 0     # the wager it was won at, and it costs nothing
        grant.remaining -= 1
        if grant.remaining <= 0:
            db.delete(grant)
        balance = available_bits(db, user)
    else:
        multiple = can_wager_multiple(user.bank_rank)
        try:
            wager = int(wager)
        except (TypeError, ValueError):
            raise Refused("Pick a whole number wager.")
        if not multiple:
            wager, capped = 1, True          # below Sovereign: 300 BITS, no multiple
        if wager < 1 or wager > MAX_WAGER:
            raise Refused("Wager must be between 1 and %s." % f"{MAX_WAGER:,}")

        cost = SPIN_COST * wager
        balance = available_bits(db, user)
        if cost > balance:
            if balance <= 0:
                raise Refused("You have no BITS yet, so there is nothing to spin with. "
                              "How BITS work: " + BITS_HELP)
            raise Refused("That costs %s BITS; you have %s." % (f"{cost:,}", f"{balance:,}"))

    seed = secrets.token_hex(32)
    spin = Spin(id=str(uuid.uuid4()), user_id=user.id, wager=wager, capped=capped, cost=cost,
                status="pending", commit=commit_of(seed), seed=seed,
                created_at=utcnow(), expires_at=utcnow() + dt.timedelta(seconds=SPIN_TTL))
    db.add(spin)
    if cost:
        db.add(Ledger(user_id=user.id, spin_id=spin.id, delta=-cost, reason="spin"))
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
        index = pick_index(spin.seed)
        slice_ = WHEEL[index]
        spin.index_landed, spin.label, spin.kind = index, slice_["label"], slice_["kind"]
        if slice_["kind"] == "bits":
            spin.payout = slice_["amount"] * spin.wager
        elif slice_["kind"] == "free":
            spin.payout, spin.prize_qty = 0, 0      # the free spins land at claim
        else:
            # a wager multiplies logos and prefixes
            spin.payout = 0
            spin.prize_qty = slice_["amount"] * spin.wager if slice_["scales"] else slice_["amount"]
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
    if spin.payout:
        db.add(Ledger(user_id=user.id, spin_id=spin.id, delta=spin.payout, reason="payout"))
    if spin.kind == "free":
        grant_free_spins(db, user.id, spin.wager, FREE_SPIN_AWARD)
        spin.prize_qty = FREE_SPIN_AWARD            # what the claim actually handed over
    spin.status = "settled"
    spin.settled_at = now
    db.commit()
    return spin, True


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


def bank_row(db, lichess_id):
    blob = meta_get(db, "bank_rows", "")
    if not blob:
        return None
    lid = name_key(lichess_id)
    for row in json.loads(blob):
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


# --------------------------------------------------------------------------- pages
CSS = """
:root{--bg:#0b0d12;--panel:#141821;--line:#242b39;--ink:#e9edf6;--muted:#93a0b8;
--gold:#f2c14e;--gold2:#b98f2e;--good:#4ade80;--bad:#f87171}
*{box-sizing:border-box}
body{margin:0;background:radial-gradient(1100px 560px at 50% -8%,#1a2130,#0b0d12 62%);
color:var(--ink);font:16px/1.55 -apple-system,BlinkMacSystemFont,'Segoe UI',Inter,sans-serif}
a{color:var(--gold)}
.wrap{max-width:1080px;margin:0 auto;padding:26px 20px 70px}
header.top{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:20px}
.brand{font-family:Georgia,'Times New Roman',serif;font-size:1.32rem}
.brand b{color:var(--gold)}
.spacer{flex:1}
.card{background:linear-gradient(180deg,rgba(255,255,255,.045),rgba(255,255,255,.015));
border:1px solid var(--line);border-radius:18px;padding:22px}
.row{display:grid;grid-template-columns:1fr 1fr;gap:20px;align-items:start}
@media(max-width:880px){.row{grid-template-columns:1fr}}
.balance{font-family:Georgia,serif;font-size:2.5rem;line-height:1.05;margin:6px 0 2px}
.lab{color:var(--muted);font-size:.76rem;text-transform:uppercase;letter-spacing:.14em}
.pill{display:inline-block;padding:4px 10px;border-radius:999px;border:1px solid var(--line);
font-size:.78rem;color:var(--muted)}
.pill.ok{border-color:rgba(74,222,128,.5);color:#9ae6b4}
.pill.no{border-color:rgba(248,113,113,.5);color:#fca5a5}
button,.btn{cursor:pointer;font:inherit;border-radius:12px;border:1px solid var(--line);
background:#1b2130;color:var(--ink);padding:10px 16px;transition:.15s}
button:hover:not(:disabled){border-color:var(--gold2);transform:translateY(-1px)}
button:disabled{opacity:.45;cursor:not-allowed}
.primary{background:linear-gradient(180deg,var(--gold),var(--gold2));color:#1a1200;font-weight:700;border:0}
.wager{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0 6px}
.wager button{padding:8px 14px;font-weight:600}
.wager button.on{border-color:var(--gold);background:rgba(242,193,78,.16)}
input{font:inherit;background:#0e131c;border:1px solid var(--line);color:var(--ink);
border-radius:10px;padding:10px 12px;width:100%}
canvas{display:block;margin:0 auto;max-width:100%}
table{width:100%;border-collapse:collapse;font-size:.9rem}
th,td{text-align:left;padding:8px 10px;border-bottom:1px solid var(--line)}
th{color:var(--muted);font-weight:600;font-size:.74rem;text-transform:uppercase;letter-spacing:.1em}
.mono{font-variant-numeric:tabular-nums}
.d-pos{color:#4ade80}.d-neg{color:#f87171}
.btnrow{display:flex;gap:12px;flex-wrap:wrap;align-items:stretch}
.bigbtn{display:inline-flex;flex-direction:column;align-items:center;justify-content:center;
gap:2px;padding:11px 26px;border-radius:14px;min-width:158px}
.bigbtn b{font-size:1.22rem;letter-spacing:.05em;line-height:1.15}
.bigbtn span{font-size:.8rem;font-weight:600;opacity:.88}
.bigbtn.free{background:linear-gradient(180deg,#7ee0a8,#38a967);color:#06220f;border:0}
.bigbtn.free span{opacity:.75}
.hint{color:var(--muted);font-size:.86rem}
.note{border:1px solid rgba(242,193,78,.38);background:rgba(242,193,78,.07);
border-radius:14px;padding:14px 18px;margin:0 0 18px;font-size:.92rem;line-height:1.5}
.err{color:#fca5a5;background:rgba(248,113,113,.08);border:1px solid rgba(248,113,113,.35);
padding:10px 14px;border-radius:12px;margin:12px 0}
.result{min-height:56px;display:flex;align-items:center;justify-content:center;
font-family:Georgia,serif;font-size:1.45rem;color:var(--gold);text-align:center}
footer{margin-top:32px;color:var(--muted);font-size:.8rem;text-align:center}
"""

# The wheel is drawn exactly like wheelofnames: pastel slices, dark text, white hub,
# a small dark pointer at the top and a soft shadow under the disc.
WHEEL_JS = r"""
const SLICES = %(slices)s, COST = %(cost)d, TTL = %(ttl)d, MULTI = %(multi)s;
const FREE0 = %(free)s;
const SEG = SLICES.length, SIZE = 560, STEP = 2 * Math.PI / SEG;
const cv = document.getElementById('wheel'), ctx = cv.getContext('2d');
const DPR = Math.min(2, window.devicePixelRatio || 1);
cv.width = SIZE * DPR; cv.height = SIZE * DPR;
cv.style.width = SIZE + 'px'; cv.style.height = SIZE + 'px';
ctx.scale(DPR, DPR);
const C = SIZE / 2, R = C - 26, HUB = R * 0.20;
let angle = -Math.PI / 2, spinning = false, spinId = null, wager = 1;

/* Emoji: use the system emoji font first, because that is what the reference wheel
   looks like. If this device has no colour emoji font, fall back to the built-in
   sheet so that nobody is ever shown an empty box. */
const sheet = new Image();
let useSheet = [], sheetReady = false;

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

useSheet = SLICES.map(s => !emojiIsColour(s.emoji));
if (useSheet.some(Boolean)) {
  sheet.onload = () => { sheetReady = true; drawWheel(); };
  sheet.src = '/emoji-sheet.png';
}

function drawWheel() {
  ctx.clearRect(0, 0, SIZE, SIZE);

  ctx.save();
  ctx.shadowColor = 'rgba(0,0,0,.5)'; ctx.shadowBlur = 24; ctx.shadowOffsetY = 7;
  ctx.beginPath(); ctx.arc(C, C, R, 0, 2 * Math.PI);
  ctx.fillStyle = '#ffffff'; ctx.fill();
  ctx.restore();

  for (let i = 0; i < SEG; i++) {
    const a0 = angle + i * STEP;
    ctx.beginPath(); ctx.moveTo(C, C); ctx.arc(C, C, R, a0, a0 + STEP); ctx.closePath();
    ctx.fillStyle = SLICES[i].color; ctx.fill();
  }
  ctx.strokeStyle = 'rgba(255,255,255,.9)'; ctx.lineWidth = 1.5;
  for (let i = 0; i < SEG; i++) {
    const a = angle + i * STEP;
    ctx.beginPath(); ctx.moveTo(C, C);
    ctx.lineTo(C + Math.cos(a) * R, C + Math.sin(a) * R); ctx.stroke();
  }

  /* label reads from the hub outwards: text, a gap, then the emoji at the rim */
  for (let i = 0; i < SEG; i++) {
    const mid = angle + (i + 0.5) * STEP;
    ctx.save(); ctx.translate(C, C); ctx.rotate(mid);
    let fs = R * 0.088, tw = 0, es = 0, gap = 0, total = 0;
    for (let k = 0; k < 6; k++) {
      es = fs * 1.02; gap = fs * 0.34;
      ctx.font = '600 ' + fs.toFixed(1) + 'px Helvetica, Arial, sans-serif';
      tw = ctx.measureText(SLICES[i].label).width;
      total = tw + gap + es;
      if (total <= R * 0.62) break;
      fs *= 0.93;
    }
    const outer = R * 0.93, textRight = outer - es - gap;
    ctx.fillStyle = '#2b2b2b';
    ctx.font = '600 ' + fs.toFixed(1) + 'px Helvetica, Arial, sans-serif';
    ctx.textAlign = 'right'; ctx.textBaseline = 'middle';
    ctx.fillText(SLICES[i].label, textRight, 0);

    const ex = outer - es / 2;
    if (useSheet[i] && sheetReady) {
      ctx.drawImage(sheet, i * 48, 0, 48, 48, ex - es / 2, -es / 2, es, es);
    } else {
      ctx.font = (es * 0.98).toFixed(1) + 'px "Apple Color Emoji","Segoe UI Emoji",' +
                 '"Noto Color Emoji",sans-serif';
      ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
      ctx.fillText(SLICES[i].emoji, ex, es * 0.04);
    }
    ctx.restore();
  }

  ctx.beginPath(); ctx.arc(C, C, HUB, 0, 2 * Math.PI);
  ctx.fillStyle = '#ffffff'; ctx.fill();
  ctx.lineWidth = 2; ctx.strokeStyle = 'rgba(0,0,0,.13)'; ctx.stroke();
  ctx.fillStyle = spinning ? '#c3c9d4' : '#2b2b2b';
  ctx.font = '700 ' + Math.round(HUB * 0.33) + 'px Helvetica, Arial, sans-serif';
  ctx.textAlign = 'center'; ctx.textBaseline = 'middle';
  ctx.fillText(spinning ? '\u2026' : 'SPIN', C, C + 2);

  /* the pointer sits on the right, exactly like the reference wheel */
  const tipX = C + R * 0.945, baseX = C + R * 1.075, hh = R * 0.088;
  ctx.beginPath();
  ctx.moveTo(tipX, C); ctx.lineTo(baseX, C - hh); ctx.lineTo(baseX, C + hh);
  ctx.closePath();
  ctx.fillStyle = '#ffffff'; ctx.fill();
  ctx.lineWidth = 2; ctx.strokeStyle = 'rgba(96,102,116,.9)'; ctx.stroke();
}

function show(msg, bad) {
  const el = document.getElementById('err');
  el.textContent = msg || '';
  el.style.display = msg ? 'block' : 'none';
  if (bad === false) {
    el.style.color = '#9ae6b4'; el.style.background = 'rgba(74,222,128,.08)';
    el.style.borderColor = 'rgba(74,222,128,.35)';
  }
}

function busy(on) {
  spinning = on;
  document.getElementById('go').disabled = on;
  document.getElementById('gofree').disabled = on;
  document.querySelectorAll('.wager button').forEach(b => { b.disabled = on || !MULTI; });
  document.getElementById('wager').disabled = on || !MULTI;
  drawWheel();
}

/* the free-spin button: how many are left, and at which wager */
function applyFree(fs) {
  const btn = document.getElementById('gofree'), label = document.getElementById('freebtn');
  const hint = document.getElementById('freehint');
  if (!fs || !fs.count) { btn.hidden = true; btn.disabled = true; hint.hidden = true; return; }
  btn.hidden = false;
  btn.disabled = spinning;
  hint.hidden = false;
  const head = fs.queue[0];
  let txt = '×' + head.count + ' left at wager ×' + head.wager.toLocaleString();
  const rest = fs.queue.slice(1);
  if (rest.length) {
    txt += ' · then ' + rest.map(q =>
      '×' + q.count + ' at ×' + q.wager.toLocaleString()).join(', ');
  }
  label.textContent = txt;
}

function setCost(w) {
  const txt = 'for ' + (COST * w).toLocaleString() + ' BITS';
  document.getElementById('costbtn').textContent = txt;
  document.getElementById('cost').textContent = (COST * w).toLocaleString();
}

function emojiFor(label) {
  const s = SLICES.filter(x => x.label === label)[0];
  return s ? s.emoji : '';
}

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

const VMAX = 0.0125;    /* radians per millisecond once it is up to speed */
const RAMP = 480;       /* ms to get there */
const GUARD = 16000;    /* absolute ceiling if the server never answers */

/* Ask where the wheel stopped. The server refuses until the wheel has really had time to
   turn, so a page closed early never learns the result and the spin becomes a forfeit. */
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

/* Collect the prize. If this never happens, nothing is paid. */
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

async function spin(useFree) {
  if (spinning) return;
  show('');
  document.getElementById('result').textContent = '';
  const res = await fetch('/api/spin', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify(useFree ? {free: true} : {wager: wager})
  });
  const j = await res.json();
  if (!res.ok) { show(j.detail || 'Refused', true); return; }
  if (j.capped) {
    show('Below Sovereign the wager is fixed at one, so this spin costs ' + COST + ' BITS.', false);
  }
  spinId = j.spin_id;
  document.getElementById('balance').textContent = j.balance.toLocaleString();
  applyFree(j.free_spins);
  document.getElementById('commit').textContent = 'committed before you paid: ' + j.commit.slice(0, 16) + '…';
  busy(true);

  /* Phase one: turn the wheel. Time drives it, so a 144 Hz screen and a 60 Hz screen see
     the same speed. The reveal is asked for as soon as the server will allow it. */
  const revealMs = Number(j.reveal_after || 3.6) * 1000;
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
    show('The server stopped answering. Nothing was collected — check Your spins.', true);
    loadLog();
    return;
  }
  const d = got;

  /* Phase two: slow down onto the slice. A cubic ease-out leaves at the speed the wheel
     already has and arrives at zero, so the last turns are the slow ones. */
  const TAU = 2 * Math.PI;
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
    document.getElementById('result').textContent = '\u{1F621} Lost — the wheel was not collected';
    show(e.message || 'The spin was not collected.', true);
    document.getElementById('commit').textContent = '';
    loadLog();
    return;
  }

  const em = emojiFor(paid.label);
  const extra = paid.free_spins && paid.free_spins.count
    ? '  \u2014 you have ' + paid.free_spins.count + ' free spin' +
      (paid.free_spins.count === 1 ? '' : 's') + ' at \u00d7' + paid.free_spins.wager.toLocaleString()
    : '';
  document.getElementById('result').textContent = paid.kind === 'free'
    ? em + ' ' + paid.label + '  \u2192  +' + paid.prize_qty.toLocaleString() + ' free spins'
    : (paid.kind === 'prize'
        ? em + ' ' + paid.label + ' \u00d7 ' + paid.prize_qty.toLocaleString()
        : em + ' ' + paid.label + '  \u2192  +' + paid.paid.toLocaleString() + ' BITS');
  if (paid.kind === 'free') document.getElementById('result').textContent += extra;
  document.getElementById('balance').textContent = paid.balance.toLocaleString();
  document.getElementById('commit').textContent =
    'wager x' + paid.wager + ' \u00b7 seed ' + paid.seed.slice(0, 12) + '\u2026 \u00b7 sha256 ' +
    paid.commit.slice(0, 12) + '\u2026';
  applyFree(paid.free_spins);
  busy(false);
  loadLog();
}

function pickWager(btn) {
  document.querySelectorAll('.wager button').forEach(b => b.classList.remove('on'));
  btn.classList.add('on');
  wager = +btn.dataset.w;
  document.getElementById('wager').value = wager;
  setCost(wager);
}

async function loadLog() {
  const rows = await (await fetch('/api/spins')).json();
  document.querySelector('#log tbody').innerHTML = rows.map(s => {
    let what;
    if (s.status === 'settled') {
      what = emojiFor(s.label) + ' ' + s.label +
        (s.kind === 'prize' ? ' \u00d7 ' + s.prize_qty.toLocaleString()
                            : '  +' + s.payout.toLocaleString());
    } else if (s.status === 'forfeited') {
      what = '\u{1F621} YOU LEFT';
    } else {
      what = 'still turning\u2026';
    }
    return '<tr><td class="mono">' + new Date(s.created_at).toLocaleString() + '</td>' +
      '<td class="mono">x' + s.wager + (s.free ? ' free' : '') + '</td><td>' + what + '</td>' +
      '<td class="mono">' + (s.cost ? '-' + s.cost.toLocaleString() : 'free') + '</td></tr>';
  }).join('') || '<tr><td colspan="4" class="hint">No spins yet.</td></tr>';
}

document.getElementById('go').onclick = function(){ spin(false); };
document.getElementById('gofree').onclick = function(){ spin(true); };
cv.onclick = function(){ spin(false); };
cv.style.cursor = 'pointer';
document.querySelectorAll('.wager button').forEach(b => { b.onclick = () => pickWager(b); });
document.getElementById('wager').oninput = e => {
  wager = Math.max(1, parseInt(e.target.value || '1', 10));
  setCost(wager);
};
applyFree(FREE0);
if (COST !== 300) {
  const note = document.getElementById('costnote');
  if (note) {
    note.innerHTML = 'Each spin costs <b>' + COST + ' \u00d7 wager</b> BITS. This one: ' +
      '<b id="cost">' + COST + '</b> BITS.';
  }
}
setCost(wager);
drawWheel();
loadLog();
"""


def esc(v):
    return (str(v or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def page(body, title="Real Humans · Mystery BITS"):
    return ("<!doctype html><html lang=en><head><meta charset=utf-8>"
            "<meta name=viewport content='width=device-width,initial-scale=1'>"
            "<title>%s</title><style>%s</style></head><body><div class=wrap>%s</div></body></html>"
            % (esc(title), CSS, body))


def login_page(message="", notice=""):
    local_card = """
  <div class=card><div class=lab>Or a local account</div>
    <p class=hint>Use this only if the Lichess button will not work for you. Your name is checked
    against the bank before you can spin, and a local account can never be admin.</p>
    <form method=post action="/auth/local" style="display:grid;gap:10px">
      <input name=username placeholder="Lichess username, no @" required>
      <input name=password type=password placeholder="Password, 8 characters or more"
        required minlength=8>
      <div style="display:flex;gap:8px">
        <button class=primary name=mode value=login>Log in</button>
        <button name=mode value=register>Register</button></div>
    </form></div>""" if ALLOW_LOCAL else """
  <div class=card><div class=lab>How the money works</div>
    <p class=hint>Every spin costs 300 BITS times your wager. Above Sovereign you may multiply the
    wager; below it the wager is fixed at one, so a spin is 300 BITS for the number on the wheel.</p>
    <p class=hint>Your BITS come from the Real Humans BITS Bank. The figure here is what the bank
    said last time the admin updated it, plus or minus everything you have done since.</p></div>"""
    top = ('<div class=err>%s</div>' % esc(message)) if message else ""
    low = ('<div class=hint>%s</div>' % esc(notice)) if notice else ""
    return page("""
<header class=top><div class=brand>Real Humans · <b>Mystery BITS</b></div></header>
%s
<div class=row>
  <div class=card><div class=lab>Sign in</div>
    <h2 style="font-family:Georgia,serif;margin:8px 0 6px">Log in with Lichess</h2>
    <p class=hint>One click. This site never sees your Lichess password; Lichess only tells it your
    name.</p>
    <p><a class=btn style="display:inline-block;text-decoration:none;padding:12px 22px"
      href="/auth/lichess">Continue with Lichess →</a></p>%s</div>%s
</div>
<footer>Spending BITS are tentative until the admin updates the BITS Bank.</footer>
""" % (top, low, local_card))


def delta_line(delta: int) -> str:
    """The sentence under the balance: where that number comes from, change coloured."""
    if delta == 0:
        return ("Your remaining BITS shown above is your BITS Bank figure, unchanged — you have "
                "not spun since the admin last updated.")
    word, cls = ("plus", "d-pos") if delta > 0 else ("minus", "d-neg")
    return ('Your remaining BITS shown above is your BITS Bank figure '
            '<b class="%s">%s %s</b> — everything you have done here since the admin last '
            'updated.' % (cls, word, f"{abs(delta):,}"))


def index_page(user, state, slices_json):
    multi = state["multiple"]
    wager_buttons = "".join(
        '<button data-w="%d"%s%s>x%s</button>' % (w, ' class=on' if w == 1 else "",
                                                  "" if multi else " disabled", f"{w:,}")
        for w in (1, 2, 5, 10, 25, 50, 100, 250, 500, 1000))
    rank_pill = ('<span class="pill ok">%s · multiples unlocked</span>' % esc(user.bank_rank)
                 if multi else '<span class="pill no">%s · %d BITS per spin, no multiples</span>'
                 % (esc(user.bank_rank or "no rank"), SPIN_COST))
    admin_link = ('<a class=btn href="/admin" style="text-decoration:none;margin-right:8px">Admin'
                  '</a>' if is_admin(user) else "")
    zero = ""
    if state["available"] <= 0:
        zero = ('<div class=note><b>You have no BITS yet.</b> Everything is visible, but a spin '
                'costs %d BITS, so the button will refuse until the bank has a figure for you. '
                '<a href="%s" target=_blank rel=noopener>How BITS work →</a> If your Lichess name '
                'is not in the BITS Bank yet, say so in the team chat and it gets added.'
                % (SPIN_COST, esc(BITS_HELP)))
    js = WHEEL_JS % {"slices": slices_json, "ttl": SPIN_TTL, "cost": SPIN_COST,
                     "multi": "true" if multi else "false",
                     "free": json.dumps(state.get("free"))}
    return page("""
<header class=top><div class=brand>Real Humans · <b>Mystery BITS</b></div>
<div class=spacer></div><span class=pill>@%s</span>%s
%s<a class=btn href="/logout" style="text-decoration:none">Log out</a></header>
%s
<div class=row>
  <div class=card>
    <div class=lab>Spending BITS · tentative until the bank is updated</div>
    <div class=balance id=balance>%s</div>
    <div class=hint>%s
      <a href="%s" target=_blank rel=noopener>How BITS work</a></div>
    <hr style="border:0;border-top:1px solid var(--line);margin:18px 0">
    <div class=lab>Wager</div>
    <div class=wager>%s</div>
    <input id=wager type=number min=1 step=1 value=1 style="max-width:160px"%s>
    <p class=hint id=costnote>Each spin costs <b>300 × wager</b> BITS. This one: <b id=cost>300</b>
    BITS.</p>
    <div id=result class=result></div>
    <div id=err class=err style="display:none"></div>
    <div class=btnrow>
      <button id=go class="bigbtn primary"><b>SPIN!</b><span id=costbtn>for 300 BITS</span></button>
      <button id=gofree class="bigbtn free" hidden><b>FREE SPIN</b><span id=freebtn></span></button>
    </div>
    <p class=hint id=freehint hidden>A free spin costs nothing, keeps the wager it was won at, and
    waits until you use it. Winning more of them adds to the pile.</p>
    <div class=hint id=commit></div>
    <p class=hint>If you reload or close the page before the wheel stops, the spin is lost — the
    wager and anything it landed on. The server enforces that, not this page.</p>
  </div>
  <div class=card style="text-align:center">
    <canvas id=wheel></canvas>
    <p class=hint>The slice is drawn on the server before you pay, and revealed when it stops.
    Click the hub to spin.</p>
  </div>
</div>
<div class=card style="margin-top:20px"><div class=lab>Your spins</div>
  <table id=log><thead><tr><th>When</th><th>Wager</th><th>Landed</th><th>Cost</th></tr></thead>
  <tbody></tbody></table></div>
<footer>Real Humans · Mystery BITS · emoji artwork from Twemoji (CC-BY 4.0) ·
  <a href="%s" target=_blank rel=noopener>BITS guide</a></footer>
<script>%s</script>
""" % (esc(user.display or user.lichess_id), rank_pill, admin_link, zero,
       f"{state['available']:,}", delta_line(state["delta"]), esc(BITS_HELP), wager_buttons,
       "" if multi else " disabled", esc(BITS_HELP), js))


PLAYER_MODAL = """
<style>
.prow{cursor:pointer}
.prow:hover td{background:rgba(242,193,78,.07)}
#pmodal{position:fixed;inset:0;z-index:60;display:none;padding:28px 16px;overflow:auto;
  background:rgba(6,8,13,.84);backdrop-filter:blur(4px)}
#pmodal.on{display:block}
.pm-card{max-width:940px;margin:0 auto;background:#11151e;border:1px solid var(--line);
  border-radius:18px;padding:22px}
.pm-head{display:flex;align-items:center;gap:12px;flex-wrap:wrap;margin-bottom:4px}
.pm-name{font-family:Georgia,serif;font-size:1.35rem}
.pm-x{margin-left:auto;padding:6px 13px;border-radius:10px}
.pm-sum{display:grid;grid-template-columns:repeat(auto-fit,minmax(118px,1fr));gap:10px;margin:16px 0 18px}
.pm-sum div{border:1px solid var(--line);border-radius:12px;padding:10px 12px}
.pm-sum span{display:block;font-size:.68rem;letter-spacing:.12em;text-transform:uppercase;color:var(--muted)}
.pm-sum b{font-size:1.08rem;font-variant-numeric:tabular-nums}
.pm-table{max-height:54vh;overflow:auto;border-top:1px solid var(--line)}
.pm-table th{position:sticky;top:0;background:#11151e}
.pm-bad{color:#fca5a5}
</style>
<div id=pmodal role=dialog aria-modal=true aria-label="Player spins">
  <div class=pm-card>
    <div class=pm-head>
      <span class=pm-name id=pmName>&mdash;</span>
      <span class=pill id=pmMeta></span>
      <button class=pm-x id=pmClose type=button>Close &times;</button>
    </div>
    <div class=pm-sum id=pmSum></div>
    <div class=pm-table>
      <table>
        <thead><tr><th>When</th><th>Wager</th><th>Landed</th><th>Status</th><th>Cost</th>
          <th>Paid / owed</th></tr></thead>
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
  async function open(uid){
    var r = await fetch('/admin/player/' + uid + '.json');
    if (!r.ok) { alert('Could not load that player.'); return; }
    var d = await r.json(), s = d.summary;
    document.getElementById('pmName').textContent = '@' + d.display;
    document.getElementById('pmMeta').textContent =
      [d.rank || 'no rank', d.in_bank ? 'in bank' : 'not in bank'].join(' \\u00b7 ');
    document.getElementById('pmSum').innerHTML = [
      ['Spins', money(s.spins)], ['Collected', money(s.collected)],
      ['Left early', money(s.forfeited)], ['Staked', money(s.staked)],
      ['Paid back', money(s.won)], ['Net', money(s.won - s.staked)],
      ['Biggest win', money(s.best)]
    ].concat(d.free ? [['Free spins', d.free.total + ' at x' + d.free.wager]] : [])
    .map(function(p){ return '<div><span>' + p[0] + '</span><b>' + p[1] + '</b></div>'; }).join('');
    document.getElementById('pmBody').innerHTML = d.spins.map(function(x){
      var what = x.status === 'settled'
        ? (x.emoji + ' ' + x.label + (x.kind === 'prize' ? ' &times; ' + money(x.prize_qty) : ''))
        : (x.status === 'forfeited' ? '\\u{1F621} left' : 'still turning');
      var paid = x.status === 'settled'
        ? (x.payout ? '+' + money(x.payout) : (x.prize_qty ? money(x.prize_qty) + ' prize' : '&mdash;'))
        : '&mdash;';
      return '<tr><td class=mono>' + new Date(x.when).toLocaleString() + '</td>' +
        '<td class=mono>x' + x.wager + '</td><td>' + esc(what) + '</td>' +
        '<td' + (x.status === 'forfeited' ? ' class=pm-bad' : '') + '>' +
          esc(x.status) + (x.free ? ' \u00b7 free' : '') + '</td>' +
        '<td class=mono>' + (x.cost ? '-' + money(x.cost) : 'free') + '</td>' +
        '<td class=mono>' + paid + '</td></tr>';
    }).join('') || '<tr><td colspan=6 class=hint>No spins yet.</td></tr>';
    pm.classList.add('on');
  }
  document.getElementById('pmClose').addEventListener('click', function(){ pm.classList.remove('on'); });
  pm.addEventListener('click', function(e){ if (e.target === pm) pm.classList.remove('on'); });
  document.addEventListener('keydown', function(e){ if (e.key === 'Escape') pm.classList.remove('on'); });
  document.querySelectorAll('tr.prow').forEach(function(tr){
    tr.addEventListener('click', function(){ open(tr.dataset.uid); });
  });
})();
</script>
"""


def nicetime(raw):
    try:
        return _aware(dt.datetime.fromisoformat(raw)).strftime("%d %b %Y, %H:%M UTC")
    except (TypeError, ValueError):
        return ""


def admin_page(rows, prizes, meta):
    body = "".join(
        "<tr class=prow data-uid=\"%d\" title=\"Click for this player's spins\">"
        "<td>@%s%s</td><td>%s</td><td class=mono>%s</td><td class=mono>%s</td>"
        "<td class=mono>%s</td><td class=mono>%s</td><td class=mono>%s</td><td>%s</td></tr>"
        % (r["id"], esc(r["display"]),
           (' <span class="pill no">\U0001F621 %d</span>' % r["forfeits"]) if r["forfeits"] else "",
           esc(r["rank"] or "-"), f"{r['bank']:,}", f"{r['delta']:+,}",
           f"{r['available']:,}", r["spins"], f"{r['staked']:,}",
           '<span class="pill ok">in bank</span>' if r["in_bank"] else
           '<span class="pill no">not in bank</span>')
        for r in rows)
    prize_rows = "".join("<tr><td>@%s</td><td>%s %s</td><td class=mono>%s</td></tr>"
                         % (esc(p["display"]), emoji_of(p["label"]), esc(p["label"]),
                            f"{p['qty']:,}")
                         for p in prizes)
    if not prize_rows:
        prize_rows = ('<tr><td colspan=3 class=hint style="padding:14px 6px">Nothing owed right now.'
                      '</td></tr>')
    cleared = nicetime(meta.get("prizes_since") or "")
    cleared_note = ("Cleared " + cleared + "." if cleared else
                    "Nothing cleared yet, so this is every prize ever won.")
    return page("""
<header class=top><div class=brand>Real Humans · <b>admin</b></div>
<div class=spacer></div><span class=pill>%s players in the bank</span>
<span class=pill>read %s</span><a class=btn href="/" style="text-decoration:none">Site</a></header>
<div class=card style="margin-bottom:18px">
  <div class=lab>Since your last update</div>
  <p class=hint>Everything you have to do is in one file: <b>Deltas + prizes</b>. Type each delta
  into column H of the sheet (that moves J, and J − P×1e6 is the spending figure), hand out the
  prizes in the last column, then <b>Sync bank</b> and <b>Mark applied</b>. Nothing is written
  for you. Mark applied sets both lists back to zero: the Deltas file and the prize obligations
  below. Someone marked <i>not in bank</i> has no BITS, so they cannot spin yet.</p>
  <form method=post action="/admin/sync" style="display:inline"><button class=primary>Sync bank</button></form>
  <form method=post action="/admin/apply" style="display:inline"
    onsubmit="return confirm('Deltas typed into the sheet and prizes handed out? This clears both lists.')">
    <button>Mark applied</button></form>
  <a class=btn href="/admin/deltas.csv" style="text-decoration:none;display:inline-block">Deltas + prizes</a>
  <a class=btn href="/admin/spins.csv" style="text-decoration:none;display:inline-block">Every spin</a>
  <a class=btn href="/admin/prizes.csv" style="text-decoration:none;display:inline-block">Prizes only</a>
  <form method=post action="/admin/upload" enctype=multipart/form-data style="margin-top:14px">
    <div class=lab>Or upload the workbook</div>
    <input type=file name=file accept=.xlsx style="margin:8px 0"><button>Upload and sync</button></form>
</div>
<div class=card style="margin-bottom:18px"><div class=lab>Prize obligations</div>
  <p class=hint>A wager of 30,000 landing on Moving Logo owes 30,000 moving logos. Only prizes won
  since the last <b>Mark applied</b> are listed: pressing it empties this card and the prizes_owed
  column of the Deltas file. <i>%s</i></p>
  <table><thead><tr><th>Player</th><th>Prize</th><th>Owed</th></tr></thead><tbody>%s</tbody></table></div>
<div class=card><div class=lab>Players</div>
  <table><thead><tr><th>Player</th><th>Rank</th><th>Bank</th><th>Delta</th><th>Available</th>
  <th>Spins</th><th>Staked</th><th>Account</th></tr></thead><tbody>%s</tbody></table></div>
<footer>Admin is the Lichess account @CC_Prestiged.</footer>
""" % (meta["bank_players"], esc(meta["bank_synced_at"] or "never"), cleared_note,
       prize_rows, body) + PLAYER_MODAL)


# --------------------------------------------------------------------------- app
app = FastAPI(docs_url=None, redoc_url=None, title="Real Humans · Mystery BITS")


def current_user(request: Request, db):
    payload = unsign(request.cookies.get(COOKIE))
    return db.get(User, payload["uid"]) if payload and "uid" in payload else None


def is_admin(user) -> bool:
    """Admin is the Lichess account in BS_ADMIN_LICHESS, and only when it is really that
    account: a password account can never be admin, whatever it calls itself."""
    if not user:
        return False
    if user.is_admin:                       # only ever set by a bank sync or by you, in the DB
        return True
    return not user.password_hash and name_key(user.lichess_id) == ADMIN_LICHESS


def state_of(db, user):
    delta = delta_for(db, user.id, user.applied_through or 0)
    return {"bank": int(user.bank_bits or 0), "delta": delta,
            "available": int(user.bank_bits or 0) + delta,
            "multiple": can_wager_multiple(user.bank_rank),
            "rank": user.bank_rank or "", "rank_name": rank_of(user.bank_rank),
            "free": free_state(db, user)}


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


@app.get("/emoji-sheet.png")
def emoji_sheet():
    blob = base64.b64decode(EMOJI_SHEET.split(",", 1)[1])
    return Response(blob, media_type="image/png",
                    headers={"Cache-Control": "public, max-age=31536000"})


@app.get("/healthz")
def healthz():
    return {"ok": True}


@app.get("/", response_class=Response)
def home(request: Request, db=Depends(get_db)):
    sweep(db)
    user = current_user(request, db)
    if not user:
        return Response(login_page(), media_type="text/html")
    slices = json.dumps([{k: v for k, v in slice_.items() if k != "weight"} for slice_ in WHEEL],
                        ensure_ascii=False)
    return Response(index_page(user, state_of(db, user), slices), media_type="text/html")


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
    return login_cookie(RedirectResponse("/"), user)


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
    return login_cookie(RedirectResponse("/"), user)


@app.get("/logout")
def logout():
    resp = RedirectResponse("/")
    resp.delete_cookie(COOKIE)
    return resp


# --- player api
@app.get("/api/me")
def api_me(request: Request, db=Depends(get_db)):
    user = current_user(request, db)
    if not user:
        return JSONResponse({"detail": "not signed in"}, status_code=401)
    s = state_of(db, user)
    s.update({"display": user.display, "cost_per_wager": SPIN_COST, "ttl": SPIN_TTL,
              "reveal_after": REVEAL_AFTER, "free_spins": free_state(db, user),
              "is_admin": is_admin(user)})
    return s


@app.post("/api/spin")
async def api_spin(request: Request, db=Depends(get_db)):
    user = current_user(request, db)
    if not user:
        return JSONResponse({"detail": "not signed in"}, status_code=401)
    try:
        body = await request.json()
    except Exception:                                     # noqa: BLE001
        body = {}
    if not isinstance(body, dict):
        body = {}
    use_free = bool(body.get("free"))
    try:
        spin, cost, balance = start_spin(db, user, body.get("wager"), free=use_free)
    except Refused as exc:
        return JSONResponse({"detail": exc.message}, status_code=400)
    return {"spin_id": spin.id, "cost": cost, "balance": balance, "commit": spin.commit,
            "wager": spin.wager, "capped": spin.capped, "ttl": SPIN_TTL,
            "free": use_free, "free_spins": free_state(db, user),
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
    return {"index": spin.index_landed, "label": spin.label, "kind": spin.kind,
            "payout": spin.payout, "prize_qty": spin.prize_qty, "wager": spin.wager,
            "seed": spin.seed, "commit": spin.commit,
            "check": {"sha256_seed": commit_of(spin.seed) == spin.commit,
                      "index_from_seed": pick_index(spin.seed)}}


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
            "balance": state_of(db, user)["available"], "replayed": not fresh,
            "free": spin.cost == 0, "free_spins": free_state(db, user),
            "seed": spin.seed, "commit": spin.commit, "status": spin.status}


@app.get("/api/spins")
def api_spins(request: Request, db=Depends(get_db)):
    user = current_user(request, db)
    if not user:
        return JSONResponse({"detail": "not signed in"}, status_code=401)
    sweep(db)
    rows = db.execute(select(Spin).where(Spin.user_id == user.id)
                      .order_by(Spin.created_at.desc()).limit(200)).scalars().all()
    return [{"id": s.id, "created_at": s.created_at.isoformat(), "wager": s.wager,
             "cost": s.cost, "status": s.status, "label": s.label or "",
             "kind": s.kind or "", "payout": s.payout or 0,
             "prize_qty": s.prize_qty or 0, "free": s.cost == 0} for s in rows]


# --- admin
def emoji_of(label):
    for slice_ in WHEEL:
        if slice_["label"] == label:
            return slice_["emoji"]
    return ""


def admin_rows(db):
    blob = meta_get(db, "bank_rows", "")
    known = {row["id"] for row in json.loads(blob)} if blob else set()
    left = dict(db.execute(select(Spin.user_id, func.count(Spin.id))
                           .where(Spin.status == "forfeited")
                           .group_by(Spin.user_id)).all())
    out = []
    for u in db.query(User).order_by(User.id).all():
        delta = delta_for(db, u.id, u.applied_through or 0)
        cnt, staked = db.execute(select(func.count(Spin.id), func.coalesce(func.sum(Spin.cost), 0))
                                 .where(Spin.user_id == u.id)).one()
        out.append({"id": u.id, "display": u.display or u.lichess_id, "rank": u.bank_rank or "",
                    "bank": int(u.bank_bits or 0), "delta": delta,
                    "available": int(u.bank_bits or 0) + delta, "spins": cnt,
                    "staked": int(staked or 0),
                    "in_bank": name_key(u.lichess_id) in known,
                    "forfeits": int(left.get(u.id, 0))})
    return out


def prizes_since(db):
    """Prizes are an obligation only until they have been handed out. 'Mark applied'
    stamps the moment, and that stamp is what empties the list."""
    raw = meta_get(db, "prizes_applied_at", "")
    try:
        return _aware(dt.datetime.fromisoformat(raw)) if raw else None
    except ValueError:
        return None


def prize_rows(db):
    since = prizes_since(db)
    q = (select(User.display, User.lichess_id, Spin.label, func.sum(Spin.prize_qty))
         .join(User, Spin.user_id == User.id)
         .where(Spin.prize_qty > 0, Spin.status == "settled", Spin.kind != "free"))
    if since:
        q = q.where(Spin.settled_at > since)
    q = q.group_by(User.id, Spin.label).order_by(func.sum(Spin.prize_qty).desc())
    return [{"display": d or l, "label": lab, "qty": int(qty)} for d, l, lab, qty in db.execute(q)]


def admin_guard(request, db):
    user = current_user(request, db)
    if is_admin(user):
        return user
    if ADMIN_KEY and hmac.compare_digest(request.headers.get("x-admin-key", ""), ADMIN_KEY):
        return user
    return None


@app.get("/admin", response_class=Response)
def admin_home(request: Request, db=Depends(get_db)):
    if not admin_guard(request, db):
        who = current_user(request, db)
        if who:
            message = ("You are signed in as @%s, but this page belongs to @%s. If that is the "
                       "wrong name, fix BS_ADMIN_LICHESS in your host's environment settings."
                       % (who.lichess_id, ADMIN_RAW.strip().lstrip("@")))
        else:
            message = "Admin is the Lichess account @%s. Log in with it." % ADMIN_RAW.strip().lstrip("@")
        return Response(login_page(message), media_type="text/html", status_code=403)
    rows_blob = meta_get(db, "bank_rows", "")
    since = prizes_since(db)
    meta = {"bank_synced_at": meta_get(db, "bank_synced_at", ""),
            "bank_players": len(json.loads(rows_blob)) if rows_blob else 0,
            "prizes_since": since.isoformat() if since else ""}
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


@app.get("/admin/player/{uid}.json")
def admin_player_json(uid: int, request: Request, db=Depends(get_db)):
    if not admin_guard(request, db):
        return JSONResponse({"detail": "forbidden"}, status_code=403)
    user = db.get(User, uid)
    if not user:
        return JSONResponse({"detail": "no such player"}, status_code=404)
    spins = db.execute(select(Spin).where(Spin.user_id == uid)
                       .order_by(Spin.created_at.desc()).limit(600)).scalars().all()
    total = db.execute(select(func.count(Spin.id))
                       .where(Spin.user_id == uid)).scalar()
    collected, staked, won = db.execute(
        select(func.count(Spin.id), func.coalesce(func.sum(Spin.cost), 0),
               func.coalesce(func.sum(Spin.payout), 0))
        .where(Spin.user_id == uid, Spin.status == "settled")).one()
    forfeited = db.execute(select(func.count(Spin.id))
                           .where(Spin.user_id == uid, Spin.status == "forfeited")).scalar()
    best = db.execute(select(func.coalesce(func.max(Spin.payout), 0))
                      .where(Spin.user_id == uid, Spin.status == "settled")).scalar()
    return {
        "display": user.display or user.lichess_id,
        "lichess_id": user.lichess_id,
        "rank": user.bank_rank or "",
        "in_bank": bank_row(db, user.lichess_id) is not None,
        "bank": int(user.bank_bits or 0),
        "delta": delta_for(db, uid, user.applied_through or 0),
        "free": free_state(db, user),
        "summary": {"spins": int(total or 0), "collected": int(collected or 0),
                    "staked": int(staked or 0), "won": int(won or 0),
                    "forfeited": int(forfeited or 0), "best": int(best or 0)},
        "spins": [{"when": x.created_at.isoformat(), "wager": x.wager, "cost": x.cost,
                   "label": x.label or "", "emoji": emoji_of(x.label), "kind": x.kind or "",
                   "payout": x.payout or 0, "prize_qty": x.prize_qty or 0,
                   "free": x.cost == 0, "status": x.status} for x in spins],
    }


@app.get("/admin/deltas.csv")
def admin_deltas(request: Request, db=Depends(get_db)):
    if not admin_guard(request, db):
        return Response("forbidden", status_code=403)
    buf = io.StringIO()
    w = csv.writer(buf)
    owed = {}
    for p in prize_rows(db):
        owed.setdefault(name_key(p["display"]), []).append(
            "%s x%s" % (p["label"], f"{p['qty']:,}"))
    w.writerow(["username", "rank", "spending_before", "delta_to_type_into_H",
                "spending_after", "spins", "staked", "prizes_owed"])
    for r in admin_rows(db):
        prizes = "; ".join(owed.get(name_key(r["display"]), []))
        if r["delta"] == 0 and r["spins"] == 0 and not prizes:
            continue
        w.writerow(["@" + str(r["display"]).lstrip("@"), r["rank"], r["bank"], r["delta"],
                    r["available"], r["spins"], r["staked"], prizes])
    return PlainTextResponse(buf.getvalue(), media_type="text/csv",
                             headers={"Content-Disposition": 'attachment; filename="rh-deltas.csv"'})


@app.get("/admin/spins.csv")
def admin_spins(request: Request, db=Depends(get_db)):
    if not admin_guard(request, db):
        return Response("forbidden", status_code=403)
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["when", "username", "wager", "capped", "cost", "free_spin", "status", "landed",
                "kind", "bits_paid", "prize_owed", "seed", "commit"])
    q = select(Spin, User).join(User, Spin.user_id == User.id).order_by(Spin.created_at)
    for spin, user in db.execute(q).all():
        w.writerow([spin.created_at.isoformat(), "@" + (user.lichess_id or ""), spin.wager,
                    "yes" if spin.capped else "", spin.cost,
                    "yes" if spin.cost == 0 else "", spin.status, spin.label or "",
                    spin.kind or "", spin.payout or 0, spin.prize_qty or 0,
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
        w.writerow(["@" + str(p["display"]).lstrip("@"), p["label"], p["qty"]])
    return PlainTextResponse(buf.getvalue(), media_type="text/csv",
                             headers={"Content-Disposition": 'attachment; filename="rh-prizes.csv"'})


# --------------------------------------------------------------------------- emoji fallback
# 16 tiles of 48 px, same order as WHEEL. Twemoji artwork, CC-BY 4.0.
EMOJI_SHEET = "data:image/png;base64,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"


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=int(os.environ.get("PORT", "8000")))
