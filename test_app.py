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


def test_unapproved_cannot_spin(db):
    user = mkuser(db, "waiting", bits=10_000, approved=False)
    with pytest.raises(main.Refused):
        main.start_spin(db, user, 1)


def test_bad_wager(db):
    user = mkuser(db, "sov", bits=10**12)
    for bad in (0, -3, main.MAX_WAGER + 1, "abc"):
        with pytest.raises(main.Refused):
            main.start_spin(db, user, bad)


def test_bits_prize_is_multiplied_by_wager(db):
    user = mkuser(db, "sov", bits=10**9)
    spin, _cost, _ = main.start_spin(db, user, 1000)
    force_index(db, spin, 11)                          # 1000 BITS
    spin, fresh = main.settle_spin(db, user, spin.id)
    assert fresh and spin.payout == 1_000_000
    assert spin.prize_qty == 0


def test_moving_logo_is_multiplied_by_wager(db):
    user = mkuser(db, "sov", bits=20_000_000)
    spin, cost, _ = main.start_spin(db, user, 30_000)
    assert cost == 9_000_000
    force_index(db, spin, 13)                          # Moving Logo
    spin, _ = main.settle_spin(db, user, spin.id)
    assert spin.label == "Moving Logo"
    assert spin.kind == "prize"
    assert spin.prize_qty == 30_000
    assert spin.payout == 0


def test_two_more_spins_is_never_multiplied(db):
    user = mkuser(db, "sov", bits=10**10)
    for wager in (1, 3, 1000, 30_000):
        spin, _c, _ = main.start_spin(db, user, wager)
        force_index(db, spin, 15)
        spin, _ = main.settle_spin(db, user, spin.id)
        assert spin.label == "2 More Spins"
        assert spin.prize_qty == 2, wager


def test_logos_and_prefixes_do_multiply(db):
    user = mkuser(db, "sov", bits=10**10)
    for index, label in ((12, "Static Logo"), (13, "Moving Logo"), (14, "Custom Prefix")):
        spin, _c, _ = main.start_spin(db, user, 50)
        force_index(db, spin, index)
        spin, _ = main.settle_spin(db, user, spin.id)
        assert (spin.label, spin.prize_qty) == (label, 50)


def test_below_sovereign_gets_exactly_one_prize(db):
    user = mkuser(db, "herald", rank="Herald", bits=10**6)
    for index in (12, 13, 14):
        spin, cost, _ = main.start_spin(db, user, 500)      # 500 is ignored, wager forced to 1
        assert cost == 300
        force_index(db, spin, index)
        spin, _ = main.settle_spin(db, user, spin.id)
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
        main.settle_spin(db, user, spin.id)
    assert db.get(main.Spin, spin.id).status == "forfeited"
    assert main.available_bits(db, user) == 700        # refunded nothing


def test_settle_twice_pays_once(db):
    user = mkuser(db, "sov", bits=1_000_000)
    spin, _c, _ = main.start_spin(db, user, 1)
    force_index(db, spin, 2)                           # 100 BITS
    main.settle_spin(db, user, spin.id)
    _spin, fresh = main.settle_spin(db, user, spin.id)
    assert fresh is False
    assert main.available_bits(db, user) == 1_000_000 - 300 + 100


def test_commit_reveal_verifies(db):
    user = mkuser(db, "sov", bits=1_000_000)
    spin, _c, _ = main.start_spin(db, user, 1)
    commit = spin.commit
    main.settle_spin(db, user, spin.id)
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
    main.settle_spin(db, winner, spin.id)
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


def test_full_spin_round_trip(db, client):
    user = mkuser(db, "player", rank="Sovereign", bits=1_000_000)
    login(client, user.id)
    assert "Sign in" not in client.get("/").text
    r = client.post("/api/spin", json={"wager": 5})
    assert r.status_code == 200, r.text
    data = r.json()
    assert data["cost"] == 1500 and data["balance"] == 998_500
    s = client.post("/api/spin/%s/settle" % data["spin_id"]).json()
    assert s["check"]["sha256_seed"] is True
    assert s["index"] == s["check"]["index_from_seed"]
    log = client.get("/api/spins").json()
    assert log[0]["status"] == "settled"


def test_capped_flag_reaches_the_api(db, client):
    user = mkuser(db, "herald", rank="Herald", bits=1_000_000)
    login(client, user.id)
    data = client.post("/api/spin", json={"wager": 50}).json()
    assert data["wager"] == 1 and data["capped"] is True and data["cost"] == 300


def test_waiting_user_sees_notice(db, client):
    user = mkuser(db, "waiting", bits=1_000, approved=False)
    login(client, user.id)
    assert "Waiting for the admin" in client.get("/").text


def test_local_register_then_login(db, client):
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


def test_local_register_of_the_admin_name_is_refused(client):
    r = client.post("/auth/local", data={"username": "CC_Prestiged", "password": "longenough",
                                         "mode": "register"})
    assert r.status_code == 403


def test_local_login_can_be_turned_off(client, monkeypatch):
    monkeypatch.setattr(main, "ALLOW_LOCAL", False)
    r = client.post("/auth/local", data={"username": "someone", "password": "longenough",
                                         "mode": "register"})
    assert r.status_code == 403
    assert "Or a local account" not in client.get("/").text


def test_admin_can_approve_a_pending_account(db, client):
    admin = mkuser(db, CC, bits=10**9)
    pending = mkuser(db, "latecomer", bits=0, approved=False)
    login(client, admin.id)
    r = client.post("/admin/approve/%d" % pending.id, follow_redirects=False)
    assert r.status_code == 303
    db.expire_all()
    assert db.get(main.User, pending.id).approved is True


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


def test_a_normal_player_is_still_held_back(db, client):
    user = mkuser(db, "newbie", bits=1, approved=False)
    login(client, user.id)
    assert "Waiting for the admin" in client.get("/").text


def test_admin_403_says_who_you_are(db, client):
    user = mkuser(db, "someone", bits=1)
    login(client, user.id)
    r = client.get("/admin")
    assert r.status_code == 403
    assert "@someone" in r.text and main.ADMIN_RAW.lstrip("@") in r.text


def test_admin_403_asks_for_a_login_when_anonymous(client):
    r = client.get("/admin")
    assert r.status_code == 403 and "Log in with it" in r.text
