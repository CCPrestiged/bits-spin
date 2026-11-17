"""
Self-test for main.py.  Run:  pytest -q

Nothing here touches Lichess or OneDrive; the workbook is built in memory.
"""
import io
import os
import tempfile

os.environ.setdefault("BS_DB_URL", "sqlite:///" + os.path.join(tempfile.mkdtemp(), "t.db"))
os.environ.setdefault("BS_SECRET_KEY", "test-secret-do-not-reuse")
os.environ.setdefault("BS_BASE_URL", "http://testserver")

import urllib.error                                               # noqa: E402
import urllib.parse                                                # noqa: E402

import openpyxl                                                    # noqa: E402
import pytest                                                      # noqa: E402
from fastapi.testclient import TestClient                          # noqa: E402

from sqlalchemy import select                                     # noqa: E402

import main                                                        # noqa: E402

CC = "cc_prestiged"


@pytest.fixture()
def db():
    main.Base.metadata.drop_all(main.engine)
    main.Base.metadata.create_all(main.engine)
    session = main.SessionLocal()
    try:
        yield session
    finally:
        session.close()


@pytest.fixture()
def client():
    return TestClient(main.app)


# --------------------------------------------------------------------- helpers
def mkuser(db, lid, rank="Sovereign", bits=10_000_000, approved=True):
    u = main.User(lichess_id=lid, display=lid, bank_rank=rank, bank_bits=bits,
                  approved=approved, is_admin=(lid == CC))
    db.add(u)
    db.commit()
    return u


def seed_for(index):
    """A 0x seed whose modulo puts the pointer exactly inside slice `index`."""
    return "%064x" % (index * 1000 + 500)


def force_index(db, spin, index):
    spin.seed = seed_for(index)
    spin.commit = main.commit_of(spin.seed)
    db.commit()
    return spin


def finish(db, user, spin, index=None):
    """reveal and collect in one go, with the clock pushed past the reveal delay"""
    later = main.utcnow() + __import__("datetime").timedelta(seconds=main.REVEAL_AFTER + 1)
    if index is not None:
        force_index(db, spin, index)
    main.reveal_spin(db, user, spin.id, now=later)
    return main.claim_spin(db, user, spin.id, now=later)


def workbook(rows, sheet=main.BANK_SHEET):
    """rows: list of (sheet_row, username, J, P, K, O)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for col, head in ((2, "Username"), (10, "Total"), (11, "Legacy Rank"), (15, "P"), (16, "Sac")):
        ws.cell(row=1, column=col, value=head)
    for row, name, total, sac, rank, prestige in rows:
        ws.cell(row=row, column=2, value="@" + name)
        ws.cell(row=row, column=10, value=total)
        ws.cell(row=row, column=11, value=rank)
        ws.cell(row=row, column=15, value=prestige)
        ws.cell(row=row, column=16, value=sac)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def login(client, user_id):
    client.cookies.set(main.COOKIE, main.sign({"uid": user_id}))
    return client


# --------------------------------------------------------------------- ranks
@pytest.mark.parametrize("raw,expected", [
    ("Bronze \U0001FA94", "Bronze"),
    ("Invincible Administrator+", "Administrator"),
    ("ConquerorEnforcer", "Enforcer"),
    ("Diamond \U0001F48E", "Diamond"),
    ("sovereign", "Sovereign"),
    ("Sentinel", "Sentinel"),
    ("-", ""),
    (None, ""),
])
def test_rank_parsing(raw, expected):
    assert main.rank_of(raw) == expected


@pytest.mark.parametrize("rank,allowed", [
    ("Sovereign", True), ("Configurator", True), ("Administrator", True),
    ("Diamond", True), ("Founder", True),
    ("Enforcer", False), ("Sentinel", False), ("Organiser", False), ("", False),
    ("ConquerorEnforcer", False),
])
def test_wager_gate(rank, allowed):
    assert main.can_wager_multiple(rank) is allowed


def test_ladder_step_order():
    assert main.rank_index("Sovereign") < main.rank_index("Administrator")
    assert main.rank_index("Organiser") == 0
    assert main.rank_index("Bronze") > main.rank_index("Administrator")


# --------------------------------------------------------------------- wheel
def test_wheel_matches_wheelofnames():
    labels = [w["label"] for w in main.WHEEL]
    assert labels == ["50 BITS", "75 BITS", "100 BITS", "150 BITS", "200 BITS", "250 BITS",
                      "300 BITS", "350 BITS", "400 BITS", "450 BITS", "500 BITS", "1000 BITS",
                      "Static Logo", "Moving Logo", "Custom Prefix", "2 More Spins"]
    colours = [w["color"] for w in main.WHEEL]
    assert colours == ["#a8e6cf", "#b2dfdb", "#81d4fa", "#4fc3f7", "#64b5f6", "#ce93d8",
                       "#ba68c8", "#f06292", "#f48fb1", "#ffb74d", "#ffd54f", "#fff176",
                       "#ce93d8", "#ffb74d", "#ba68c8", "#90caf9"]


def test_weight_total():
    assert main.total_weight() == 16_000


def test_pick_index_is_deterministic():
    seed = seed_for(7)
    assert main.pick_index(seed) == main.pick_index(seed) == 7


def test_every_slice_reachable():
    assert {main.pick_index(seed_for(i)) for i in range(16)} == set(range(16))


def test_commit_is_sha256_of_seed():
    import hashlib
    seed = seed_for(3)
    assert main.commit_of(seed) == hashlib.sha256(seed.encode()).hexdigest()


# --------------------------------------------------------------------- bank
def test_spending_bits_is_j_minus_p_times_a_million(db):
    blob = workbook([(12, "CC_Prestiged", 20_176_025_450, 6552, "Sovereign", 26)])
    rows = main.parse_bank(blob)
    assert len(rows) == 1
    row = rows[0]
    assert row["total"] == 20_176_025_450
    assert row["sacrificed"] == 6552
    assert row["bits"] == 20_176_025_450 - 6552 * 1_000_000 == 13_624_025_450
    assert row["rank"] == "Sovereign"


def test_parse_skips_blank_and_header_rows():
    blob = workbook([(2, "Alpha", 1_000_000, 0, "Herald", 0),
                     (3, "", 0, 0, "", 0)])
    assert [r["id"] for r in main.parse_bank(blob)] == ["alpha"]


def test_parse_rejects_wrong_sheet():
    with pytest.raises(main.BankError):
        main.parse_bank(workbook([], sheet="Something Else"))


def test_sync_bank_updates_only_known_users(db):
    mkuser(db, "alpha", bits=0)
    main.sync_bank(db, blob=workbook([(2, "Alpha", 9_000_000, 1, "Sovereign", 1),
                                      (3, "Stranger", 5, 0, "Herald", 0)]), origin="test")
    alpha = db.query(main.User).filter_by(lichess_id="alpha").one()
    assert alpha.bank_bits == 8_000_000
    assert alpha.approved is True
    assert db.query(main.User).count() == 1


# --------------------------------------------------------------------- engine
def test_below_sovereign_is_forced_to_one(db):
    user = mkuser(db, "herald", rank="Herald", bits=1_000_000)
    spin, cost, balance = main.start_spin(db, user, 40)
    assert spin.wager == 1 and spin.capped is True
    assert cost == 300
    assert balance == 999_700


def test_sovereign_keeps_the_multiple(db):
    user = mkuser(db, "sov", rank="Sovereign", bits=1_000_000)
    spin, cost, _ = main.start_spin(db, user, 1000)
    assert spin.wager == 1000 and spin.capped is False
    assert cost == 300_000


def test_not_enough_bits(db):
    user = mkuser(db, "poor", bits=500)
    with pytest.raises(main.Refused):
        main.start_spin(db, user, 2)
    main.start_spin(db, user, 1)                      # 300 is fine


def test_no_approval_needed_but_the_wallet_is_the_gate(db):
    user = mkuser(db, "waiting", bits=10_000, approved=False)
    main.start_spin(db, user, 1)                      # never approved, still spins
    poor = mkuser(db, "broke", bits=0, approved=True)
    with pytest.raises(main.Refused) as err:
        main.start_spin(db, poor, 1)
    assert "no BITS yet" in str(err.value) and main.BITS_HELP in str(err.value)


def test_bad_wager(db):
    user = mkuser(db, "sov", bits=10**12)
    for bad in (0, -3, main.MAX_WAGER + 1, "abc"):
        with pytest.raises(main.Refused):
            main.start_spin(db, user, bad)


def test_bits_prize_is_multiplied_by_wager(db):
    user = mkuser(db, "sov", bits=10**9)
    spin, _cost, _ = main.start_spin(db, user, 1000)
    force_index(db, spin, 11)                          # 1000 BITS
    spin, fresh = finish(db, user, spin)
    assert fresh and spin.payout == 1_000_000
    assert spin.prize_qty == 0


def test_moving_logo_is_multiplied_by_wager(db):
    user = mkuser(db, "sov", bits=20_000_000)
    spin, cost, _ = main.start_spin(db, user, 30_000)
    assert cost == 9_000_000
    force_index(db, spin, 13)                          # Moving Logo
    spin, _ = finish(db, user, spin)
    assert spin.label == "Moving Logo"
    assert spin.kind == "prize"
    assert spin.prize_qty == 30_000
    assert spin.payout == 0


def test_two_more_spins_always_hands_over_two(db):
    user = mkuser(db, "sov", bits=10**10)
    for wager in (1, 3, 1000, 30_000):
        spin, _c, _ = main.start_spin(db, user, wager)
        force_index(db, spin, 15)
        spin, _ = finish(db, user, spin)
        assert spin.label == "2 More Spins"
        assert spin.prize_qty == 2, wager
        assert spin.payout == 0
        assert main.free_state(db, user)["wager"] == wager
        # clear the queue between rounds so each round starts fresh
        for grant in main.free_grants(db, user.id):
            db.delete(grant)
        db.commit()


def test_logos_and_prefixes_do_multiply(db):
    user = mkuser(db, "sov", bits=10**10)
    for index, label in ((12, "Static Logo"), (13, "Moving Logo"), (14, "Custom Prefix")):
        spin, _c, _ = main.start_spin(db, user, 50)
        force_index(db, spin, index)
        spin, _ = finish(db, user, spin)
        assert (spin.label, spin.prize_qty) == (label, 50)


def test_below_sovereign_gets_exactly_one_prize(db):
    user = mkuser(db, "herald", rank="Herald", bits=10**6)
    for index in (12, 13, 14):
        spin, cost, _ = main.start_spin(db, user, 500)      # 500 is ignored, wager forced to 1
        assert cost == 300
        force_index(db, spin, index)
        spin, _ = finish(db, user, spin)
        assert spin.prize_qty == 1


def test_scales_flags_are_explicit():
    flags = {w["label"]: w["scales"] for w in main.WHEEL}
    assert flags["2 More Spins"] is False
    assert all(flags[k] for k in ("Static Logo", "Moving Logo", "Custom Prefix", "1000 BITS"))


def test_expired_spin_is_forfeited(db):
    user = mkuser(db, "sov", bits=1_000)
    spin, _c, after = main.start_spin(db, user, 1)
    assert after == 700
    spin.expires_at = main.utcnow() - __import__("datetime").timedelta(seconds=1)
    db.commit()
    with pytest.raises(main.Refused):
        main.reveal_spin(db, user, spin.id)
    assert db.get(main.Spin, spin.id).status == "forfeited"
    assert main.available_bits(db, user) == 700        # refunded nothing


def test_claiming_twice_pays_once(db):
    user = mkuser(db, "sov", bits=1_000_000)
    spin, _c, _ = main.start_spin(db, user, 1)
    force_index(db, spin, 2)                           # 100 BITS
    finish(db, user, spin)
    later = main.utcnow() + __import__("datetime").timedelta(seconds=main.REVEAL_AFTER + 2)
    _spin, fresh = main.claim_spin(db, user, spin.id, now=later)
    assert fresh is False
    assert main.available_bits(db, user) == 1_000_000 - 300 + 100


def test_commit_reveal_verifies(db):
    user = mkuser(db, "sov", bits=1_000_000)
    spin, _c, _ = main.start_spin(db, user, 1)
    commit = spin.commit
    finish(db, user, spin)
    assert main.commit_of(spin.seed) == commit
    assert main.pick_index(spin.seed) == spin.index_landed


def test_balance_reflects_ledger(db):
    user = mkuser(db, "sov", bits=1_000)
    main.start_spin(db, user, 1)
    assert main.available_bits(db, user) == 700
    db.add(main.Ledger(user_id=user.id, spin_id="x", delta=5_000, reason="manual"))
    db.commit()
    assert main.available_bits(db, user) == 5_700


def test_apply_marks_everything_as_banked(db):
    user = mkuser(db, "sov", bits=1_000)
    main.start_spin(db, user, 1)
    last = db.query(main.Ledger).order_by(main.Ledger.id.desc()).first().id
    user.applied_through = last
    db.commit()
    assert main.delta_for(db, user.id, user.applied_through) == 0


# --------------------------------------------------------------------- admin
def test_admin_is_cc_prestiged(db):
    assert main.is_admin(mkuser(db, CC, bits=1))
    assert main.is_admin(mkuser(db, "someone", bits=1)) is False
    assert main.is_admin(None) is False


def test_admin_page_blocks_players(db, client):
    user = mkuser(db, "someone", bits=1_000_000)
    login(client, user.id)
    assert client.get("/admin").status_code == 403
    assert client.get("/admin/spins.csv").status_code == 403
    assert client.get("/admin/prizes.csv").status_code == 403


def test_admin_sees_prize_obligation(db, client):
    admin = mkuser(db, CC, bits=10**9)
    winner = mkuser(db, "winner", bits=10**9)
    spin, _c, _ = main.start_spin(db, winner, 30_000)
    force_index(db, spin, 13)
    finish(db, winner, spin)
    login(client, admin.id)
    page = client.get("/admin").text
    assert "Moving Logo" in page and "30,000" in page
    csv_text = client.get("/admin/prizes.csv").text
    assert csv_text.splitlines()[0] == "username,prize,owed"
    assert "Moving Logo,30000" in csv_text
    spins_csv = client.get("/admin/spins.csv").text
    assert "prize_owed" in spins_csv and ",30000," in spins_csv


def test_deltas_csv_uses_sheet_column(db, client):
    admin = mkuser(db, CC, bits=10**9)
    player = mkuser(db, "player", bits=1_000_000)
    main.start_spin(db, player, 1)
    login(client, admin.id)
    text = client.get("/admin/deltas.csv").text
    assert "delta_to_type_into_H" in text
    assert "player" in text and ",-300," in text


def test_admin_upload_workbook(db, client):
    admin = mkuser(db, CC, bits=0)
    mkuser(db, "alpha", bits=0, approved=False)
    login(client, admin.id)
    blob = workbook([(2, "alpha", 4_000_000, 2, "Sovereign", 3)])
    r = client.post("/admin/upload", files={"file": ("BITS Bank.xlsx", blob,
                                                     "application/vnd.openxmlformats-officedocument"
                                                     ".spreadsheetml.sheet")})
    assert r.status_code == 200
    alpha = db.query(main.User).filter_by(lichess_id="alpha").one()
    assert alpha.bank_bits == 2_000_000 and alpha.approved is True


# --------------------------------------------------------------------- http
def test_anonymous_gets_login_page(client):
    r = client.get("/")
    assert r.status_code == 200 and "Continue with Lichess" in r.text


def test_api_needs_a_session(client):
    assert client.get("/api/me").status_code == 401
    assert client.post("/api/spin", json={"wager": 1}).status_code == 401


def test_full_spin_round_trip(db, client, monkeypatch):
    monkeypatch.setattr(main, "REVEAL_AFTER", 0)        # no waiting in a unit test
    user = mkuser(db, "player", rank="Sovereign", bits=1_000_000)
    login(client, user.id)
    assert "Sign in" not in client.get("/").text
    r = client.post("/api/spin", json={"wager": 5})
    assert r.status_code == 200, r.text
    data = r.json()
    assert data["cost"] == 1500 and data["balance"] == 998_500
    assert data["reveal_after"] == 0
    rev = client.post("/api/spin/%s/reveal" % data["spin_id"]).json()
    assert rev["check"]["sha256_seed"] is True
    assert rev["index"] == rev["check"]["index_from_seed"]
    assert main.Spin is not None
    assert db.execute(select(main.Spin.status)).scalars().all() == ["revealed"]
    paid = client.post("/api/spin/%s/claim" % data["spin_id"]).json()
    assert paid["status"] == "settled"
    assert paid["paid"] == rev["payout"]
    log = client.get("/api/spins").json()
    assert log[0]["status"] == "settled"


def test_capped_flag_reaches_the_api(db, client):
    user = mkuser(db, "herald", rank="Herald", bits=1_000_000)
    login(client, user.id)
    data = client.post("/api/spin", json={"wager": 50}).json()
    assert data["wager"] == 1 and data["capped"] is True and data["cost"] == 300


def test_new_account_sees_the_wheel_and_the_help_link(db, client):
    user = mkuser(db, "waiting", bits=0, approved=False)
    login(client, user.id)
    page = client.get("/").text
    assert "id=wheel" in page and "Waiting for the admin" not in page
    assert "You have no BITS yet" in page and main.BITS_HELP in page


def test_local_login_is_off_by_default(client):
    assert main.ALLOW_LOCAL is False
    r = client.post("/auth/local", data={"username": "newbie", "password": "longenough",
                                         "mode": "register"})
    assert r.status_code == 403


def test_local_register_then_login(db, client, monkeypatch):
    monkeypatch.setattr(main, "ALLOW_LOCAL", True)
    r = client.post("/auth/local", data={"username": "newbie", "password": "longenough",
                                         "mode": "register"}, follow_redirects=False)
    assert r.status_code in (303, 307)
    client.cookies.clear()
    r = client.post("/auth/local", data={"username": "newbie", "password": "wrongpass",
                                         "mode": "login"})
    assert r.status_code == 401
    r = client.post("/auth/local", data={"username": "newbie", "password": "longenough",
                                         "mode": "login"}, follow_redirects=False)
    assert r.status_code in (303, 307)


def test_forged_cookie_is_rejected(client):
    client.cookies.set(main.COOKIE, "eyJ1aWQiOjF9.deadbeef")
    assert client.get("/api/me").status_code == 401


def test_healthz(client):
    assert client.get("/healthz").json() == {"ok": True}


# --------------------------------------------------------------------- security
def test_postgres_urls_get_a_driver():
    assert main.normalize_db_url("postgres://u:p@h/db") == "postgresql+psycopg2://u:p@h/db"
    assert main.normalize_db_url("postgresql://u@h/db") == "postgresql+psycopg2://u@h/db"
    assert main.normalize_db_url("sqlite:///x.db") == "sqlite:///x.db"


def test_password_account_named_cc_is_not_admin(db):
    fake = main.User(lichess_id=CC, display="CC_Prestiged", password_hash="pbkdf2$1$a$b")
    db.add(fake)
    db.commit()
    assert main.is_admin(fake) is False


def test_local_register_of_the_admin_name_is_refused(client, monkeypatch):
    monkeypatch.setattr(main, "ALLOW_LOCAL", True)
    r = client.post("/auth/local", data={"username": "CC_Prestiged", "password": "longenough",
                                         "mode": "register"})
    assert r.status_code == 403
    assert "signed in with Lichess" in r.text


def test_login_page_hides_the_password_form_when_off(client):
    assert main.ALLOW_LOCAL is False
    assert "Or a local account" not in client.get("/").text
    assert "Continue with Lichess" in client.get("/").text


def test_approval_route_is_gone(db, client):
    admin = mkuser(db, CC, bits=10**9)
    other = mkuser(db, "latecomer", bits=0)
    login(client, admin.id)
    assert client.post("/admin/approve/%d" % other.id, follow_redirects=False).status_code == 404


def test_sacrifice_rebuilt_when_column_p_has_no_cached_value():
    # column P is =SUMPRODUCT(...), so a file saved by a script can come back empty
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = main.BANK_SHEET
    ws.cell(row=2, column=2, value="@Somebody")
    ws.cell(row=2, column=10, value=20_176_025_450)
    ws.cell(row=2, column=15, value=26)
    ws.cell(row=2, column=16, value=None)
    buf = io.BytesIO()
    wb.save(buf)
    row = main.parse_bank(buf.getvalue())[0]
    assert row["sacrificed"] == 6552                       # 26·27·28/3
    assert row["bits"] == 20_176_025_450 - 6_552_000_000


# --------------------------------------------------------------------- lichess oauth
def test_authorize_url_is_oauth_with_pkce():
    url = main.lichess_url("st4te", "CHALLENGE")
    assert url.startswith("https://lichess.org/oauth?")
    q = urllib.parse.parse_qs(url.split("?", 1)[1])
    assert q["response_type"] == ["code"]
    assert q["code_challenge_method"] == ["S256"]
    assert q["state"] == ["st4te"] and q["client_id"] == [main.LICHESS_CLIENT_ID]


def test_token_endpoint_is_api_token(monkeypatch):
    seen = {}

    class Ok:
        def __enter__(self): return self
        def __exit__(self, *a): return False
        def read(self): return b'{"access_token": "tok", "token_type": "Bearer"}'

    def fake(req, timeout=None):
        seen["url"] = req.full_url
        seen["body"] = urllib.parse.parse_qs(req.data.decode())
        return Ok()

    monkeypatch.setattr(main.urllib.request, "urlopen", fake)
    out = main.lichess_token("thecode", "theverifier")
    assert seen["url"] == "https://lichess.org/api/token"
    assert "oauth" not in seen["url"].replace("lichess.org/api/token", "")
    assert seen["body"]["grant_type"] == ["authorization_code"]
    assert seen["body"]["code"] == ["thecode"]
    assert seen["body"]["code_verifier"] == ["theverifier"]
    assert seen["body"]["client_id"] == [main.LICHESS_CLIENT_ID]
    assert "client_secret" not in seen["body"]          # public client, no secret
    assert out["access_token"] == "tok"


def test_no_registration_credentials_are_needed():
    assert main.LICHESS_AUTHORIZE == "https://lichess.org/oauth"
    assert main.LICHESS_TOKEN == "https://lichess.org/api/token"
    assert not hasattr(main, "LICHESS_CLIENT_SECRET")
    assert main.LICHESS_CLIENT_ID and isinstance(main.LICHESS_CLIENT_ID, str)


def test_token_errors_are_explained(monkeypatch):
    body = io.BytesIO(b'{"error":"invalid_grant","error_description":"code expired"}')

    def boom(req, timeout=None):
        raise urllib.error.HTTPError(req.full_url, 400, "Bad Request", {}, body)

    monkeypatch.setattr(main.urllib.request, "urlopen", boom)
    with pytest.raises(main.LichessError) as err:
        main.lichess_token("x", "y" * 43)
    assert "code expired" in str(err.value)


def test_authorize_redirect_carries_pkce(client):
    r = client.get("/auth/lichess", follow_redirects=False)
    assert r.status_code in (302, 303, 307)
    loc = r.headers["location"]
    assert loc.startswith("https://lichess.org/oauth?")
    assert "code_challenge_method=S256" in loc and "code_challenge=" in loc
    assert "state=" in loc and "redirect_uri=" in loc
    assert "bs_oauth" in r.cookies or "bs_oauth" in client.cookies


def test_callback_signs_the_player_in(db, client, monkeypatch):
    monkeypatch.setattr(main, "lichess_token", lambda code, verifier: {"access_token": "tok"})
    monkeypatch.setattr(main, "lichess_account",
                        lambda token: {"id": "alpha", "username": "Alpha"})
    client.cookies.set("bs_oauth", main.sign({"state": "st", "verifier": "v" * 43}))
    r = client.get("/auth/callback?code=abc&state=st", follow_redirects=False)
    assert r.status_code in (302, 303, 307)
    user = db.query(main.User).filter_by(lichess_id="alpha").one()
    assert user.display == "Alpha"
    assert client.get("/api/me").status_code == 200


def test_callback_rejects_a_wrong_state(db, client, monkeypatch):
    monkeypatch.setattr(main, "lichess_token", lambda c, v: {"access_token": "tok"})
    client.cookies.set("bs_oauth", main.sign({"state": "one", "verifier": "v" * 43}))
    r = client.get("/auth/callback?code=abc&state=two")
    assert r.status_code == 400
    assert db.query(main.User).count() == 0


# --------------------------------------------------------------------- admin casing
def test_name_key_folds_case_space_and_at():
    assert main.name_key("CC_Prestiged") == main.name_key("  @cc_Prestiged ") == "cc_prestiged"
    assert main.name_key(None) == "" and main.name_key("") == ""


def test_admin_survives_a_messy_env_value(db, monkeypatch):
    monkeypatch.setattr(main, "ADMIN_LICHESS", main.name_key("  @CC_Prestiged\n"))
    user = mkuser(db, "cc_prestiged", bits=1)
    assert main.is_admin(user) is True


def test_admin_with_any_casing_wins(db, monkeypatch):
    monkeypatch.setattr(main, "ADMIN_LICHESS", main.name_key("CC_PRESTIGED"))
    user = mkuser(db, "CC_Prestiged", bits=1)
    assert main.is_admin(user) is True


def test_first_login_before_any_bank_upload(db, client, monkeypatch):
    """The catch-22: nobody is approved until the bank is uploaded, and only the admin can
    upload it. The admin must be usable on the very first login."""
    monkeypatch.setattr(main, "ADMIN_LICHESS", main.name_key("CC_Prestiged"))
    monkeypatch.setattr(main, "lichess_token", lambda c, v: {"access_token": "t"})
    monkeypatch.setattr(main, "lichess_account",
                        lambda t: {"id": "cc_prestiged", "username": "CC_Prestiged"})
    client.cookies.set("bs_oauth", main.sign({"state": "st", "verifier": "v" * 43}))
    assert client.get("/auth/callback?code=c&state=st", follow_redirects=False).status_code in (302, 303, 307)

    user = db.query(main.User).filter_by(lichess_id="cc_prestiged").one()
    assert user.approved is True and user.is_admin is True
    assert "Waiting for the admin" not in client.get("/").text
    assert client.get("/admin").status_code == 200


def test_a_pending_admin_still_gets_the_wheel(db, client):
    user = mkuser(db, CC, bits=1, approved=False)
    login(client, user.id)
    page = client.get("/").text
    assert "Waiting for the admin" not in page
    assert "id=wheel" in page


def test_a_player_with_bits_sees_no_warning(db, client):
    user = mkuser(db, "rich", bits=5_000, approved=False)
    login(client, user.id)
    page = client.get("/").text
    assert "You have no BITS yet" not in page and "id=wheel" in page


def test_admin_403_says_who_you_are(db, client):
    user = mkuser(db, "someone", bits=1)
    login(client, user.id)
    r = client.get("/admin")
    assert r.status_code == 403
    assert "@someone" in r.text and main.ADMIN_RAW.lstrip("@") in r.text


def test_admin_403_asks_for_a_login_when_anonymous(client):
    r = client.get("/admin")
    assert r.status_code == 403 and "Log in with it" in r.text


# --------------------------------------------------------------------- the one worklist
def test_deltas_csv_carries_the_prizes(db, client):
    admin = mkuser(db, CC, bits=10**9)
    player = mkuser(db, "winner", bits=10**8)
    for index, wager in ((13, 2), (13, 1), (12, 4)):
        spin, _c, _ = main.start_spin(db, player, wager)
        force_index(db, spin, index)
        finish(db, player, spin)
    login(client, admin.id)
    text = client.get("/admin/deltas.csv").text
    head, row = text.splitlines()[0], text.splitlines()[1]
    assert head == ("username,rank,spending_before,delta_to_type_into_H,spending_after,"
                    "spins,staked,prizes_owed")
    assert row.startswith("@winner,")
    assert "Moving Logo x3" in row and "Static Logo x4" in row
    assert row.count(";") == 1


def test_prize_csv_and_admin_table_keep_emoji_out_of_the_csv(db, client):
    admin = mkuser(db, CC, bits=10**9)
    player = mkuser(db, "winner", bits=10**8)
    spin, _c, _ = main.start_spin(db, player, 2)
    force_index(db, spin, 13)
    finish(db, player, spin)
    login(client, admin.id)
    csv_text = client.get("/admin/prizes.csv").text
    assert "Moving Logo,2" in csv_text and "\U0001F39E" not in csv_text
    page = client.get("/admin").text
    assert "\U0001F39E" in page and "Moving Logo" in page


def test_admin_sees_who_is_not_in_the_bank(db, client):
    admin = mkuser(db, CC, bits=10**9)
    mkuser(db, "outsider", bits=0)
    main.sync_bank(db, blob=workbook([(2, "alpha", 1_000_000, 0, "Sovereign", 0)]), origin="t")
    login(client, admin.id)
    page = client.get("/admin").text
    assert "not in bank" in page and "in bank" in page


# --------------------------------------------------------------------- leaving mid-spin
def test_reveal_waits_for_the_wheel(db):
    user = mkuser(db, "sov", bits=10**6)
    spin, _c, _ = main.start_spin(db, user, 1)
    with pytest.raises(main.Refused) as err:
        main.reveal_spin(db, user, spin.id)
    assert "still turning" in str(err.value)
    assert db.get(main.Spin, spin.id).status == "pending"


def test_claim_before_reveal_is_refused(db):
    user = mkuser(db, "sov", bits=10**6)
    spin, _c, _ = main.start_spin(db, user, 1)
    with pytest.raises(main.Refused) as err:
        main.claim_spin(db, user, spin.id)
    assert "has not stopped" in str(err.value)


def test_closing_the_page_before_the_reveal_loses_the_wager(db):
    user = mkuser(db, "sov", bits=1_000)
    spin, _c, after = main.start_spin(db, user, 1)
    assert after == 700
    spin.expires_at = main.utcnow() - __import__("datetime").timedelta(seconds=1)
    db.commit()
    main.sweep(db)
    assert db.get(main.Spin, spin.id).status == "forfeited"
    assert main.available_bits(db, user) == 700            # the 300 is gone, nothing came back


def test_closing_the_page_after_the_reveal_loses_the_prize(db):
    user = mkuser(db, "sov", bits=10**7)
    spin, _c, _ = main.start_spin(db, user, 2)
    force_index(db, spin, 13)                              # Moving Logo x2
    later = main.utcnow() + __import__("datetime").timedelta(seconds=main.REVEAL_AFTER + 1)
    main.reveal_spin(db, user, spin.id, now=later)
    assert db.get(main.Spin, spin.id).status == "revealed"
    assert db.get(main.Spin, spin.id).prize_qty == 2

    spin = db.get(main.Spin, spin.id)
    spin.expires_at = main.utcnow() - __import__("datetime").timedelta(seconds=1)
    db.commit()
    main.sweep(db)
    assert db.get(main.Spin, spin.id).status == "forfeited"   # seen, not collected
    assert main.available_bits(db, user) == 10**7 - 600
    assert db.query(main.Ledger).filter_by(reason="payout").count() == 0
    with pytest.raises(main.Refused):
        main.claim_spin(db, user, spin.id)


def test_a_forfeited_prize_is_not_owed_to_the_admin(db, client):
    admin = mkuser(db, CC, bits=10**9)
    player = mkuser(db, "quitter", bits=10**7)
    spin, _c, _ = main.start_spin(db, player, 1)
    force_index(db, spin, 12)                              # Static Logo
    later = main.utcnow() + __import__("datetime").timedelta(seconds=main.REVEAL_AFTER + 1)
    main.reveal_spin(db, player, spin.id, now=later)
    spin = db.get(main.Spin, spin.id)
    spin.expires_at = main.utcnow() - __import__("datetime").timedelta(seconds=1)
    db.commit()
    main.sweep(db)
    login(client, admin.id)
    assert "Static Logo" not in client.get("/admin/prizes.csv").text


def test_forfeit_shows_up_as_you_left(db, client, monkeypatch):
    monkeypatch.setattr(main, "REVEAL_AFTER", 0)
    user = mkuser(db, "quitter", bits=1_000)
    spin, _c, _ = main.start_spin(db, user, 1)
    spin.expires_at = main.utcnow() - __import__("datetime").timedelta(seconds=1)
    db.commit()
    login(client, user.id)
    rows = client.get("/api/spins").json()
    assert rows[0]["status"] == "forfeited" and rows[0]["cost"] == 300


# --------------------------------------------------------------------- admin popup
def test_admin_player_popup_feed(db, client):
    admin = mkuser(db, CC, bits=10**9)
    player = mkuser(db, "watched", bits=10**8)
    won, _c, _ = main.start_spin(db, player, 3)
    force_index(db, won, 11)                               # 1000 BITS x3
    finish(db, player, won)
    lost, _c, _ = main.start_spin(db, player, 1)
    force_index(db, lost, 2)
    later = main.utcnow() + __import__("datetime").timedelta(seconds=main.REVEAL_AFTER + 1)
    main.reveal_spin(db, player, lost.id, now=later)
    lost = db.get(main.Spin, lost.id)
    lost.expires_at = main.utcnow() - __import__("datetime").timedelta(seconds=1)
    db.commit()
    main.sweep(db)

    login(client, admin.id)
    d = client.get("/admin/player/%d.json" % player.id).json()
    assert d["display"] == "watched" and d["summary"]["spins"] == 2
    assert d["summary"]["collected"] == 1
    assert d["summary"]["forfeited"] == 1
    assert d["summary"]["best"] == 3000
    statuses = sorted(x["status"] for x in d["spins"])
    assert statuses == ["forfeited", "settled"]
    assert any(x["emoji"] == "\U0001F451" for x in d["spins"])


def test_admin_player_popup_is_admin_only(db, client):
    player = mkuser(db, "watched", bits=10**6)
    assert client.get("/admin/player/%d.json" % player.id).status_code == 403
    login(client, player.id)
    assert client.get("/admin/player/%d.json" % player.id).status_code == 403
    assert client.get("/admin/player/9999.json").status_code == 403


def test_admin_page_has_clickable_rows_and_the_overlay(db, client):
    admin = mkuser(db, CC, bits=10**9)
    mkuser(db, "watched", bits=10**6)
    login(client, admin.id)
    page = client.get("/admin").text
    assert 'class=prow data-uid=' in page
    assert 'id=pmodal' in page and "pmBody" in page


# --------------------------------------------------------------------- the sentence
def test_delta_line_wording_and_colour():
    assert "minus 225" in main.delta_line(-225) and "d-neg" in main.delta_line(-225)
    assert "plus 1,440" in main.delta_line(1440) and "d-pos" in main.delta_line(1440)
    assert "unchanged" in main.delta_line(0) and "<b" not in main.delta_line(0)


def test_balance_sentence_on_the_page(db, client, monkeypatch):
    monkeypatch.setattr(main, "REVEAL_AFTER", 0)
    user = mkuser(db, "sov", bits=1_000)
    main.start_spin(db, user, 1)                          # -300 in the ledger
    login(client, user.id)
    page = client.get("/").text
    assert "Your remaining BITS shown above is your BITS Bank figure" in page
    assert '<b class="d-neg">minus 300</b>' in page


# --------------------------------------------------------------------- free spins
def win_free(db, user, wager=5):
    spin, _c, _ = main.start_spin(db, user, wager)
    force_index(db, spin, 15)                      # 2 More Spins
    return finish(db, user, spin)


def test_free_spins_are_credited_at_the_wager_that_won_them(db):
    user = mkuser(db, "sov", bits=10**7)
    spin, _ = win_free(db, user, 5)
    assert spin.kind == "free" and spin.prize_qty == 2 and spin.payout == 0
    fs = main.free_state(db, user)
    assert fs["count"] == 2 and fs["wager"] == 5


def test_free_spins_stack(db):
    user = mkuser(db, "sov", bits=10**7)
    win_free(db, user, 5)
    win_free(db, user, 5)
    assert main.free_state(db, user)["count"] == 4


def test_free_spins_queue_by_age_when_wagers_differ(db):
    user = mkuser(db, "sov", bits=10**9)
    win_free(db, user, 5)
    win_free(db, user, 50)
    fs = main.free_state(db, user)
    assert fs["queue"] == [{"wager": 5, "count": 2}, {"wager": 50, "count": 2}]
    assert fs["total"] == 4 and fs["wager"] == 5
    for _ in range(2):
        spin, cost, _ = main.start_spin(db, user, 1, free=True)
        assert cost == 0 and spin.wager == 5
    spin, _c, _ = main.start_spin(db, user, 1, free=True)
    assert spin.wager == 50                       # next in the queue
    assert main.free_state(db, user)["count"] == 1


def test_a_free_spin_costs_nothing_and_ignores_the_chosen_wager(db):
    user = mkuser(db, "sov", bits=1_000_000)
    win_free(db, user, 5)
    before = main.available_bits(db, user)
    spin, cost, after = main.start_spin(db, user, 50, free=True)
    assert cost == 0 and after == before
    assert spin.wager == 5 and spin.cost == 0
    assert main.free_state(db, user)["count"] == 1


def test_a_free_spin_pays_at_its_wager(db):
    user = mkuser(db, "sov", bits=1_000_000)
    win_free(db, user, 5)
    base = main.available_bits(db, user)
    spin, _c, _ = main.start_spin(db, user, 1, free=True)
    force_index(db, spin, 11)                     # 1000 BITS
    spin, _ = finish(db, user, spin)
    assert spin.payout == 5_000                    # 1000 x the wager it was won at
    assert main.available_bits(db, user) == base + 5_000


def test_a_free_spin_can_win_more_free_spins(db):
    user = mkuser(db, "sov", bits=1_000_000)
    win_free(db, user, 5)                          # 2 at x5
    spin, _c, _ = main.start_spin(db, user, 1, free=True)
    force_index(db, spin, 15)                      # another 2 More Spins
    finish(db, user, spin)
    assert main.free_state(db, user)["count"] == 3  # one spent, two won back


def test_no_free_spins_left_is_refused(db):
    user = mkuser(db, "sov", bits=10**6)
    with pytest.raises(main.Refused) as err:
        main.start_spin(db, user, 1, free=True)
    assert "no free spins" in str(err.value).lower()


def test_abandoning_a_free_spin_burns_it(db):
    user = mkuser(db, "sov", bits=1_000_000)
    win_free(db, user, 5)
    spin, _c, _ = main.start_spin(db, user, 1, free=True)
    assert main.free_state(db, user)["count"] == 1
    spin.expires_at = main.utcnow() - __import__("datetime").timedelta(seconds=1)
    db.commit()
    main.sweep(db)
    assert db.get(main.Spin, spin.id).status == "forfeited"
    assert main.free_state(db, user)["count"] == 1   # spent, not refunded


def test_free_spins_survive_a_reload(db):
    user = mkuser(db, "sov", bits=10**7)
    win_free(db, user, 5)
    db.expire_all()                                  # everything back from the database
    reloaded = db.get(main.User, user.id)
    assert main.free_state(db, reloaded)["count"] == 2


def test_free_spins_are_not_a_prize_to_hand_out(db, client):
    admin = mkuser(db, CC, bits=10**9)
    player = mkuser(db, "spinner", bits=10**7)
    win_free(db, player, 5)
    login(client, admin.id)
    assert "2 More Spins" not in client.get("/admin/prizes.csv").text
    d = client.get("/admin/player/%d.json" % player.id).json()
    assert d["free"]["count"] == 2 and d["free"]["wager"] == 5
    assert d["spins"][0]["free"] is False and d["spins"][0]["kind"] == "free"


def test_api_reports_free_spins(db, client, monkeypatch):
    monkeypatch.setattr(main, "REVEAL_AFTER", 0)
    user = mkuser(db, "sov", bits=10**7)
    win_free(db, user, 5)
    login(client, user.id)
    assert client.get("/api/me").json()["free"]["count"] == 2
    r = client.post("/api/spin", json={"free": True})
    assert r.status_code == 200, r.text
    d = r.json()
    assert d["cost"] == 0 and d["wager"] == 5 and d["free"] is True
    assert d["free_spins"]["count"] == 1
    client.post("/api/spin/%s/reveal" % d["spin_id"])
    paid = client.post("/api/spin/%s/claim" % d["spin_id"]).json()
    assert paid["free"] is True


def test_page_offers_the_free_button_markup(db, client):
    user = mkuser(db, "sov", bits=10**7)
    win_free(db, user, 5)
    login(client, user.id)
    page = client.get("/").text
    assert "SPIN!" in page and "id=costbtn" in page
    assert "FREE SPIN" in page and "id=freebtn" in page
    assert '"count": 2' in page and '"wager": 5' in page
