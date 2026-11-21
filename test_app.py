"""Tests for the three-wheel site.

Run with:  sh dev.sh          (installs the pins, then runs this file)

Nothing here talks to the network. Every spin is driven through the same functions the
routes call, with a seed forced to a known slice, so "where it landed" is never a guess.
"""
import io
import os
import tempfile
import uuid

os.environ["BS_DB_URL"] = "sqlite:///" + os.path.join(tempfile.mkdtemp(), "test.db")
os.environ.setdefault("BS_SECRET_KEY", "test-key")
os.environ.setdefault("BS_BASE_URL", "http://testserver")
os.environ.setdefault("BS_MAX_WAGER", "1000000")

import datetime as dt                                              # noqa: E402
import urllib.error                                                # noqa: E402
import urllib.parse                                                 # noqa: E402

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
def mkuser(db, lid, rank="Sovereign", bits=10_000_000, approved=True, title="", admin=None):
    u = main.User(lichess_id=lid, display=lid, bank_rank=rank, bank_bits=bits,
                  approved=approved, is_admin=(lid == CC) if admin is None else admin)
    db.add(u)
    db.commit()
    return u


def seed_for(wheel, index):
    """A seed whose modulo puts the pointer exactly at the start of slice `index`."""
    low = sum(int(s["weight"]) for s in main.slices(wheel)[:index])
    return "%064x" % low


def force_index(db, spin, index):
    spin.seed = seed_for(int(spin.wheel or 1), index)
    spin.commit = main.commit_of(spin.seed)
    db.commit()
    return spin


def finish(db, user, spin, index=None, now=None):
    """reveal and collect in one go, with the clock pushed past the reveal delay"""
    later = now or (main.utcnow() + dt.timedelta(seconds=main.REVEAL_AFTER + 1))
    if index is not None:
        force_index(db, spin, index)
    main.reveal_spin(db, user, spin.id, now=later)
    return main.claim_spin(db, user, spin.id, now=later)


def spin_at(db, user, wheel, index, wager=1, free=False):
    """The whole path, straight to the result: start, reveal, claim."""
    spin, cost, balance = main.start_spin(db, user, wager, free=free, wheel=wheel)
    spin, fresh = finish(db, user, spin, index=index)
    return spin, cost, balance


def index_of(wheel, label):
    for i, s in enumerate(main.slices(wheel)):
        if s["label"] == label:
            return i
    raise AssertionError("no slice called %r on wheel %s" % (label, wheel))


def workbook(rows, sheet=main.BANK_SHEET):
    """rows: list of (sheet_row, username, J, P, K, O, L)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for col, head in ((2, "Username"), (10, "Total"), (11, "Legacy Rank"), (12, "Title"),
                      (15, "P"), (16, "Sac")):
        ws.cell(row=1, column=col, value=head)
    for row in rows:
        name, total, sac, rank = row[1], row[2], row[3], row[4]
        prestige = row[5] if len(row) > 5 else 0
        title = row[6] if len(row) > 6 else ""
        ws.cell(row=row[0], column=2, value="@" + name)
        ws.cell(row=row[0], column=10, value=total)
        ws.cell(row=row[0], column=11, value=rank)
        ws.cell(row=row[0], column=12, value=title)
        ws.cell(row=row[0], column=15, value=prestige)
        ws.cell(row=row[0], column=16, value=sac)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def login(client, user_id):
    client.cookies.set(main.COOKIE, main.sign({"uid": user_id}))
    return client


class Clock:
    """A utcnow() you can move by hand, so 'before' and 'after' are not a race."""

    def __init__(self, start):
        self.t = start

    def __call__(self):
        return self.t

    def tick(self, seconds):
        self.t = self.t + dt.timedelta(seconds=seconds)
        return self.t


# ===================================================================== the wheels
@pytest.mark.parametrize("wheel", [1, 2, 3])
def test_every_wheel_adds_up_to_one_hundred_percent(wheel):
    assert main.bp_total(wheel) == 10000


def test_wheel_one_is_the_original_sixteen_slices():
    assert len(main.WHEEL_1) == 16
    assert {s["weight"] for s in main.WHEEL_1} == {625}          # 6.25% each


WHEEL_2_ODDS = {
    "Bust - 100 BITS": 45.00, "600 BITS": 15.00, "1,500 BITS": 15.00, "3,000 BITS": 10.00,
    "6,000 BITS": 6.00, "15,000 BITS": 4.00, "60,000 BITS": 0.45, "150,000 BITS": 0.05,
    "5 More Spins": 0.50, "Fragment of x2 BITS Booster": 2.20, "Fragment of x3 BITS Booster": 1.00,
    "Fragment of x5 BITS Booster": 0.60, "Fragment of x10 BITS Booster": 0.20,
}

WHEEL_3_ODDS = {
    "Bust - 0 BITS": 60.00, "1,500 BITS": 10.00, "6,000 BITS": 8.00, "15,000 BITS": 7.00,
    "30,000 BITS": 5.00, "90,000 BITS": 4.00, "300,000 BITS": 2.00,
    "3,000,000 BITS": 0.20, "Progressive Legacy Jackpot": 0.05, "1 Legacy Token": 1.10,
    "3 Legacy Tokens": 0.50, "5 Legacy Tokens": 0.15, "x3 BITS Booster": 0.60,
    "x5 BITS Booster": 0.30, "x10 BITS Booster": 0.20, "x20 BITS Booster": 0.10,
    "x3 Legacy Booster": 0.60, "x5 Legacy Booster": 0.20,
}


@pytest.mark.parametrize("wheel,expected", [(2, WHEEL_2_ODDS), (3, WHEEL_3_ODDS)])
def test_probabilities_are_the_ones_in_the_sheets(wheel, expected):
    got = {s["label"]: int(s["weight"]) / 100.0 for s in main.slices(wheel)}
    assert got == expected
    assert round(sum(expected.values()), 6) == 100.0


def test_no_wheel_has_a_duplicate_label():
    for wheel in (1, 2, 3):
        labels = [s["label"] for s in main.slices(wheel)]
        assert len(labels) == len(set(labels))


@pytest.mark.parametrize("wheel", [1, 2, 3])
def test_the_whole_number_line_lands_where_the_weights_say(wheel):
    """Every one of the 10,000 residues stops in the slice that owns it."""
    seen = {}
    for n in range(10000):
        index = main.pick_index(wheel, "%064x" % n)
        seen[index] = seen.get(index, 0) + 1
    for i, s in enumerate(main.slices(wheel)):
        assert seen.get(i, 0) == int(s["weight"]), (wheel, s["label"])


@pytest.mark.parametrize("wheel", [1, 2, 3])
def test_pick_index_is_deterministic_and_reaches_every_slice(wheel):
    seed = "a" * 64
    assert main.pick_index(wheel, seed) == main.pick_index(wheel, seed)
    reached = {main.pick_index(wheel, seed_for(wheel, i)) for i in range(len(main.slices(wheel)))}
    assert reached == set(range(len(main.slices(wheel))))


def test_commit_is_sha256_of_the_seed():
    import hashlib
    seed = "0123456789abcdef" * 4
    assert main.commit_of(seed) == hashlib.sha256(seed.encode()).hexdigest()


@pytest.mark.parametrize("wheel,per,cost,pct", [
    (1, 239.06, 300, 79.7), (2, 1965.0, 3000, 65.5), (3, 18780.0, 30000, 62.6)])
def test_average_return_is_the_bits_a_spin_hands_back(wheel, per, cost, pct):
    avg = main.average_return(wheel)
    assert round(avg["per_spin"], 2) == per
    assert avg["cost"] == cost
    assert round(avg["pct"], 1) == pct


def test_jackpot_is_not_counted_as_a_slice_value():
    """The pot is other people's money, so it must not inflate the average."""
    with_pot = main.average_return(3)["per_spin"]
    assert with_pot == 18780.0


@pytest.mark.parametrize("wheel", [1, 2, 3])
def test_every_emoji_has_a_fallback_tile(wheel):
    for s in main.slices(wheel):
        assert s["code"] in main.EMOJI_CODES
        assert s["emoji"]


@pytest.mark.parametrize("wheel", [1, 2, 3])
def test_every_slice_has_two_short_lines_for_the_wheel(wheel):
    for s in main.slices(wheel):
        assert 1 <= len(s["lines"]) <= 2


def test_items_and_fragments_line_up():
    for key in main.PRIZE_KEYS:
        assert key in main.ITEMS
        assert main.ITEMS[key]["column"]
    for frag, booster in main.FRAGMENTS.items():
        assert booster in main.ITEMS
        assert frag.startswith("frag_")
    assert len(main.PRIZE_KEYS) == 11                   # H to R


def test_the_three_sheets_columns_come_first():
    assert main.SHEET_PRIZE_COLUMNS == ["static_logo", "moving_logo", "custom_prefix"]
    assert main.PRIZE_KEYS[:3] == main.SHEET_PRIZE_COLUMNS


# ===================================================================== the ladders
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


def test_the_ladder_is_the_one_in_the_legacy_bank_sheet():
    assert main.LEGACY_LADDER[0] == "Organiser"
    assert main.LEGACY_LADDER[-1] == "Founder"
    assert main.LEGACY_LADDER.index("Sovereign") < main.LEGACY_LADDER.index("Bronze")
    assert main.TITLE_LADDER[0] == "Aspirant"
    assert main.TITLE_LADDER[-1] == "Netherfounder"


def test_rank_sorts_by_the_ladder_not_by_the_alphabet():
    organiser = main.rank_ordinal("Organiser")
    enforcer = main.rank_ordinal("Enforcer")
    sovereign = main.rank_ordinal("Sovereign")
    bronze = main.rank_ordinal("Bronze")
    founder = main.rank_ordinal("Founder")
    assert organiser < enforcer < sovereign < bronze < founder
    # alphabetically Bronze would come before Organiser, which is the bug this fixes
    assert "Bronze" < "Organiser" and bronze > organiser


def test_a_higher_title_breaks_a_tie_on_rank():
    assert main.rank_ordinal("Bronze", "Immortal Disciple") > main.rank_ordinal("Bronze")
    assert (main.rank_ordinal("Bronze", "Immortal Disciple")
            < main.rank_ordinal("Silver", "Aspirant"))


def test_no_rank_sorts_last():
    assert main.rank_ordinal("") < main.rank_ordinal("Organiser")
    assert main.rank_ordinal("nonsense") < main.rank_ordinal("Organiser")


@pytest.mark.parametrize("rank,allowed", [
    ("Sovereign", True), ("Configurator", True), ("Administrator", True),
    ("Diamond", True), ("Founder", True),
    ("Enforcer", False), ("Sentinel", False), ("Organiser", False), ("", False),
    ("ConquerorEnforcer", False),
])
def test_wheel_one_takes_a_wager_from_sovereign(rank, allowed):
    assert main.can_wager_multiple(rank, 1) is allowed


@pytest.mark.parametrize("rank,allowed", [
    ("Administrator", True), ("Administrator+", True), ("Bronze", True), ("Silver", True),
    ("Gold", True), ("Platinum", True), ("Founder", True),
    ("Invincible Administrator+", True),
    ("Configurator", False), ("Sovereign", False), ("Enforcer", False),
    ("Herald", False), ("", False),
])
def test_wheel_two_takes_a_wager_from_administrator(rank, allowed):
    assert main.can_wager_multiple(rank, 2) is allowed


def test_the_two_wheels_ask_for_different_rungs():
    assert main.wager_from(1) == "Sovereign"
    assert main.wager_from(2) == "Administrator"
    # a Sovereign may wager on Wheel I and not on Wheel II
    assert main.can_wager_multiple("Sovereign", 1) is True
    assert main.can_wager_multiple("Sovereign", 2) is False


# ===================================================================== bank
def test_spending_bits_is_j_minus_p_times_a_million(db):
    rows = main.parse_bank(workbook([(3, "someone", 13_300_000_000, 6_552, "Platinum", 26)]))
    assert rows[0]["bits"] == 13_300_000_000 - 6_552 * 1_000_000
    assert rows[0]["rank"] == "Platinum"
    assert rows[0]["prestige"] == 26


def test_the_bank_also_carries_the_hierarchical_title(db):
    rows = main.parse_bank(workbook([(3, "someone", 1000, 0, "Silver", 0, "Celestial Strategist")]))
    assert rows[0]["title"] == "Celestial Strategist"


def test_parse_skips_blank_and_header_rows():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = main.BANK_SHEET
    ws.cell(row=1, column=2, value="Username")          # the real sheet's own heading row
    ws.cell(row=2, column=2, value="")
    ws.cell(row=3, column=2, value="not a name")        # no @, so it is not a player
    ws.cell(row=4, column=2, value="@real")
    ws.cell(row=4, column=10, value=5000)
    ws.cell(row=4, column=11, value="Herald")
    buf = io.BytesIO()
    wb.save(buf)
    rows = main.parse_bank(buf.getvalue())
    assert [r["display"] for r in rows] == ["real"]


def test_parse_rejects_wrong_sheet():
    with pytest.raises(main.BankError):
        main.parse_bank(workbook([(3, "x", 1, 0, "Herald", 0)], sheet="Nope"))


def test_sync_bank_updates_only_known_users(db):
    known = mkuser(db, "known", rank="Herald", bits=0)
    unknown = mkuser(db, "unknown", rank="Herald", bits=0)
    main.sync_bank(db, blob=workbook([(3, "known", 9_000_000, 0, "Sentinel", 0)]))
    db.refresh(known)
    db.refresh(unknown)
    assert known.bank_bits == 9_000_000 and known.bank_rank == "Sentinel"
    assert unknown.bank_bits == 0


# ===================================================================== costs and wagers
@pytest.mark.parametrize("wheel,cost", [(1, 300), (2, 3000), (3, 30_000)])
def test_each_wheel_has_its_own_price(db, wheel, cost):
    user = mkuser(db, "u")
    user.test_unlocked = True
    db.commit()
    spin, paid, _ = main.start_spin(db, user, 1, wheel=wheel)
    assert paid == cost and cost == main.wheel_cost(wheel)


@pytest.mark.parametrize("wheel,max_wager", [(1, 1_000), (2, 100), (3, 10)])
def test_the_wager_range_belongs_to_the_wheel(wheel, max_wager):
    assert main.wheel_max_wager(wheel) == max_wager


def test_wheel_one_takes_the_whole_range(db):
    user = mkuser(db, "u", bits=10 ** 12)
    wager_ready(db, user, 1)
    spin, cost, _ = main.start_spin(db, user, 1_000, wheel=1)
    assert cost == 300 * 1_000
    with pytest.raises(main.Refused) as exc:
        main.start_spin(db, user, 1_001, wheel=1)
    assert "1 to 1,000" in exc.value.message


def test_wheel_two_stops_at_a_hundred(db):
    user = mkuser(db, "u", bits=10 ** 9, rank="Administrator")
    user.test_unlocked = True
    db.commit()
    wager_ready(db, user, 2)
    spin, cost, _ = main.start_spin(db, user, 100, wheel=2)
    assert spin.wager == 100 and cost == 300_000
    with pytest.raises(main.Refused) as exc:
        main.start_spin(db, user, 101, wheel=2)
    assert "1 to 100" in exc.value.message


def test_wheel_three_takes_a_wager_of_ten(db):
    user = mkuser(db, "u", bits=10 ** 9, rank="Silver")
    user.test_unlocked = True
    db.commit()
    wager_ready(db, user, 3)
    spin, cost, _ = main.start_spin(db, user, 10, wheel=3)
    assert spin.wager == 10 and cost == 300_000
    with pytest.raises(main.Refused) as exc:
        main.start_spin(db, user, 11, wheel=3)
    assert "1 to 10" in exc.value.message


@pytest.mark.parametrize("wheel", [1, 2, 3])
def test_below_the_rank_the_wager_is_forced_to_one(db, wheel):
    """The spins are all there; the rank is not."""
    user = mkuser(db, "low", rank="Enforcer")
    user.test_unlocked = True
    db.commit()
    wager_ready(db, user, wheel)
    spin, cost, _ = main.start_spin(db, user, 5, wheel=wheel)
    assert spin.wager == 1 and spin.capped is True
    assert cost == main.wheel_cost(wheel)


@pytest.mark.parametrize("wheel,after", [(1, 10), (2, 30), (3, 150)])
def test_the_wager_waits_for_the_spin_count(db, wheel, after):
    """The other half of the key: the rank is right, the finished spins are not."""
    user = mkuser(db, "u", rank="Silver", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    grant_spins(db, user, wheel, after - 1)
    spin, _, _ = main.start_spin(db, user, 5, wheel=wheel)
    assert spin.wager == 1 and spin.capped is True
    grant_spins(db, user, wheel, 1)
    spin, cost, _ = main.start_spin(db, user, 2, wheel=wheel)
    assert spin.wager == 2 and spin.capped is False
    assert cost == 2 * main.wheel_cost(wheel)


def test_asking_for_one_is_never_a_capped_spin(db):
    """A fresh player at wager one is not told off for something they did not ask for."""
    user = mkuser(db, "u")
    spin, _, _ = main.start_spin(db, user, 1, wheel=1)
    assert spin.wager == 1 and spin.capped is False


def test_not_enough_bits_is_refused(db):
    user = mkuser(db, "poor", bits=299)
    with pytest.raises(main.Refused) as exc:
        main.start_spin(db, user, 1, wheel=1)
    assert "you have" in exc.value.message


def test_no_bits_at_all_gets_the_helpful_message(db):
    user = mkuser(db, "empty", bits=0)
    with pytest.raises(main.Refused) as exc:
        main.start_spin(db, user, 1, wheel=1)
    assert "How BITS work" in exc.value.message


def test_a_bad_wager_is_refused(db):
    user = mkuser(db, "u")
    with pytest.raises(main.Refused):
        main.start_spin(db, user, "wat", wheel=1)


# ===================================================================== unlocking
def _spin_n(db, user, wheel, n, index=0):
    for _ in range(n):
        main.start_spin(db, user, 1, wheel=wheel)
        spin = db.execute(select(main.Spin).where(main.Spin.user_id == user.id,
                                                 main.Spin.status == "pending")).scalars().first()
        finish(db, user, spin, index=index)


def grant_spins(db, user, wheel, n):
    """Finished spins without the ceremony. Every count in the site reads rows with status
    settled, so a row is enough to move a bar."""
    for _ in range(n):
        db.add(main.Spin(id=uuid.uuid4().hex, user_id=user.id, wheel=int(wheel), wager=1,
                         cost=0, status="settled", created_at=main.utcnow(),
                         settled_at=main.utcnow(), seed="00" * 32, commit="00" * 32,
                         kind="bits", index_landed=0, label="", payout=0))
    db.commit()


def wager_ready(db, user, wheel):
    """The spin half of the gate; the rank half comes from the account."""
    grant_spins(db, user, wheel, main.WAGER_AFTER[int(wheel)])


def test_wheel_two_opens_on_the_fiftieth_finished_spin(db):
    user = mkuser(db, "u", bits=10 ** 9)
    grant_spins(db, user, 1, main.UNLOCK_SPINS[2] - 1)
    assert main.unlocked_wheels(db, user) == {1}
    grant_spins(db, user, 1, 1)
    assert main.unlocked_wheels(db, user) == {1, 2}


def test_a_forfeit_does_not_count_towards_unlocking(db):
    user = mkuser(db, "u", bits=10 ** 9)
    grant_spins(db, user, 1, main.UNLOCK_SPINS[2] - 1)
    main.start_spin(db, user, 1, wheel=1)
    spin = db.execute(select(main.Spin).where(main.Spin.user_id == user.id,
                                             main.Spin.status == "pending")).scalars().first()
    later = main.utcnow() + dt.timedelta(seconds=main.SPIN_TTL + 5)
    main.sweep(db, now=later)
    assert main.wheel_counts(db, user.id)[1] == main.UNLOCK_SPINS[2] - 1
    assert 2 not in main.unlocked_wheels(db, user)


def test_wheel_three_opens_after_two_hundred_wheel_two_spins(db):
    user = mkuser(db, "u", bits=10 ** 12)
    grant_spins(db, user, 1, main.UNLOCK_SPINS[2])      # Wheel II first
    assert main.unlocked_wheels(db, user) == {1, 2}
    grant_spins(db, user, 2, main.UNLOCK_SPINS[3] - 1)
    assert main.unlocked_wheels(db, user) == {1, 2}
    grant_spins(db, user, 2, 1)
    assert main.unlocked_wheels(db, user) == {1, 2, 3}


def test_a_locked_wheel_cannot_be_spun(db):
    user = mkuser(db, "u", bits=10 ** 9)
    with pytest.raises(main.Refused) as exc:
        main.start_spin(db, user, 1, wheel=2)
    assert "locked" in exc.value.message
    with pytest.raises(main.Refused):
        main.start_spin(db, user, 1, wheel=3)


def test_the_test_switch_opens_everything(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    assert main.unlocked_wheels(db, user) == {1, 2, 3}
    spin, cost, _ = main.start_spin(db, user, 1, wheel=3)
    assert cost == 30_000


def test_counts_are_kept_per_wheel(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    _spin_n(db, user, 1, 3)
    _spin_n(db, user, 2, 2)
    _spin_n(db, user, 3, 1)
    assert main.wheel_counts(db, user.id) == {1: 3, 2: 2, 3: 1}


def test_progress_bars_point_at_the_next_wheel(db):
    user = mkuser(db, "u")
    grant_spins(db, user, 1, 36)
    bar = main.progress_for(db, user, 1)
    assert bar["have"] == 36 and bar["need"] == 50 and bar["unit"] == "spins"
    assert bar["opens"] == "Wheel II" and bar["secret"] is False and bar["open"] is False
    assert main.progress_for(db, user, 2)["opens"] == "Wheel III"
    assert main.progress_for(db, user, 2)["need"] == 200
    third = main.progress_for(db, user, 3)
    assert third["secret"] is True and third["need"] == main.SECRET_TARGET == 1_000
    assert third["unit"] == "spins"


@pytest.mark.parametrize("wheel,after,need", [(1, 10, 50), (2, 30, 200), (3, 150, 1_000)])
def test_each_bar_carries_both_marks(wheel, after, need):
    """Two things wait on every bar: the wager part way along, and the end of it."""
    bar = main.progress_for(None, None, wheel, counts={wheel: 0}, have={1})
    assert [m["at"] for m in bar["marks"]] == [after, need]
    assert bar["marks"][0]["kind"] == "wager"
    assert bar["marks"][0]["label"].endswith(main.WAGER_MIN_RANK[wheel])
    assert bar["need"] == need and bar["marks"][-1]["at"] == need


def test_the_third_bar_counts_wheel_three_spins_not_bits(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin_at(db, user, 3, index_of(3, "90,000 BITS"))          # a big win
    assert main.progress_for(db, user, 3)["have"] == 1        # one spin, not 90,000
    spin_at(db, user, 3, index_of(3, "Bust - 0 BITS"))
    assert main.progress_for(db, user, 3)["have"] == 2


def test_mark_applied_leaves_every_bar_alone(db, client):
    """Only the deltas and the prize obligations are cleared. The bars count lifetime spins:
    a bar that went backwards would tell a player their spins had been taken away."""
    admin = mkuser(db, CC, bits=10 ** 12, admin=True)
    user = mkuser(db, "u", bits=10 ** 12)
    _spin_n(db, user, 1, 100)                    # Wheel II is now open
    login(client, admin.id)
    client.post("/admin/apply")
    db.expire_all()

    bar = main.progress_for(db, user, 1)
    assert bar["have"] == 100                    # exactly what it was
    assert bar["open"] is True
    assert 2 in main.unlocked_wheels(db, user)
    assert not hasattr(main.User, "progress_reset_at")


def test_the_bar_counts_every_finished_spin_it_ever_had(db):
    user = mkuser(db, "u", bits=10 ** 12)
    _spin_n(db, user, 1, 3)
    assert main.progress_for(db, user, 1)["have"] == 3
    _spin_n(db, user, 1, 2)
    assert main.progress_for(db, user, 1)["have"] == 5        # nothing takes them back
    assert main.wheel_counts(db, user.id)[1] == 5
    assert main.progress_for(db, user, 2)["have"] == 0        # each bar counts its own wheel


# ===================================================================== payouts
def test_a_bits_prize_is_multiplied_by_the_wager(db):
    user = mkuser(db, "u")
    wager_ready(db, user, 1)
    spin, cost, _ = spin_at(db, user, 1, index_of(1, "500 BITS"), wager=4)
    assert cost == 1200 and spin.payout == 2000


def test_a_wheel_two_prize_is_multiplied_by_the_wager(db):
    user = mkuser(db, "u", bits=10 ** 9, rank="Administrator")
    user.test_unlocked = True
    db.commit()
    wager_ready(db, user, 2)
    spin, cost, _ = spin_at(db, user, 2, index_of(2, "1,500 BITS"), wager=3)
    assert cost == 9000 and spin.payout == 4500


def test_the_bust_on_wheel_two_still_pays_one_hundred(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin, cost, _ = spin_at(db, user, 2, index_of(2, "Bust - 100 BITS"))
    assert cost == 3000 and spin.payout == 100


def test_the_bust_on_wheel_three_pays_nothing(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin, cost, _ = spin_at(db, user, 3, index_of(3, "Bust - 0 BITS"))
    assert spin.payout == 0


def test_logos_are_multiplied_by_the_wager_on_wheel_one(db):
    user = mkuser(db, "u")
    wager_ready(db, user, 1)
    spin, cost, _ = spin_at(db, user, 1, index_of(1, "Static Logo"), wager=30)
    assert spin.prize_qty == 30
    assert main.award_totals(db, user.id)[user.id]["static_logo"] == 30


def test_a_wager_multiplies_fragments(db):
    user = mkuser(db, "u", bits=10 ** 9, rank="Administrator")
    user.test_unlocked = True
    db.commit()
    wager_ready(db, user, 2)
    spin, cost, _ = spin_at(db, user, 2, index_of(2, "Fragment of x2 BITS Booster"), wager=10)
    assert spin.prize_qty == 10
    # ten fragments is three Boosters and one fragment towards the next
    assert main.award_totals(db, user.id)[user.id]["bits_x2"] == 3
    assert main.fragments_of(db, user.id) == {"frag_x2": 1}


def test_a_wager_multiplies_a_finished_booster_up_to_its_cap(db):
    """x3 and x5 BITS Boosters have no cap here; the x20 stops at three a spin."""
    user = mkuser(db, "u", bits=10 ** 9, rank="Silver")
    user.test_unlocked = True
    db.commit()
    wager_ready(db, user, 3)
    spin, _, _ = spin_at(db, user, 3, index_of(3, "x3 BITS Booster"), wager=10)
    assert spin.prize_qty == 10                     # nothing caps this one
    capped, _, _ = spin_at(db, user, 3, index_of(3, "x20 BITS Booster"), wager=10)
    assert capped.prize_qty == 3
    assert main.award_totals(db, user.id)[user.id]["bits_x20"] == 3


def test_every_item_on_every_wheel_multiplies_with_the_wager():
    for wheel in (1, 2, 3):
        for slice_ in main.slices(wheel):
            if slice_["kind"] == "item":
                assert slice_["scales"] is True, (wheel, slice_["label"])
    one_free = [s for s in main.slices(1) if s["kind"] == "free"]
    two_free = [s for s in main.slices(2) if s["kind"] == "free"]
    assert one_free[0]["scales"] is True and one_free[0]["cap"] == 10
    assert two_free[0]["scales"] is True and two_free[0]["cap"] == 50


def test_a_wager_multiplies_tokens(db):
    user = mkuser(db, "u", bits=10 ** 9)
    main.give_award(db, user.id, "legacy_token", 3 * 4)
    db.commit()
    assert main.award_totals(db, user.id)[user.id]["legacy_token"] == 12


def test_a_wager_multiplies_free_spins_up_to_fifty(db):
    """Wheel II's five grow with the wager and stop at fifty a spin: a x10 and a x100 both
    leave with the same fifty."""
    user = mkuser(db, "u", bits=10 ** 9, rank="Administrator")
    user.test_unlocked = True
    db.commit()
    wager_ready(db, user, 2)
    spin, cost, _ = spin_at(db, user, 2, index_of(2, "5 More Spins"), wager=10)
    assert spin.prize_qty == 50 and cost == 30_000
    assert main.free_state(db, user, 2)["count"] == 50
    spin, _, _ = spin_at(db, user, 2, index_of(2, "5 More Spins"), wager=100)
    assert spin.prize_qty == 50                     # the cap, not 500


def test_a_fragment_slice_stops_at_twenty_a_spin(db):
    user = mkuser(db, "u", bits=10 ** 9, rank="Administrator")
    user.test_unlocked = True
    db.commit()
    wager_ready(db, user, 2)
    spin, cost, _ = spin_at(db, user, 2, index_of(2, "Fragment of x2 BITS Booster"), wager=100)
    assert cost == 300_000 and spin.prize_qty == 20
    assert main.award_totals(db, user.id)[user.id]["bits_x2"] == 6    # twenty is six boosters
    assert main.fragments_of(db, user.id) == {"frag_x2": 2}


def test_the_legacy_boosters_stop_at_three_a_spin(db):
    user = mkuser(db, "u", bits=10 ** 9, rank="Silver")
    user.test_unlocked = True
    db.commit()
    wager_ready(db, user, 3)
    spin, _, _ = spin_at(db, user, 3, index_of(3, "x5 Legacy Booster"), wager=10)
    assert spin.prize_qty == 3


def test_tokens_are_not_capped_on_wheel_three(db):
    user = mkuser(db, "u", bits=10 ** 9, rank="Silver")
    user.test_unlocked = True
    db.commit()
    wager_ready(db, user, 3)
    spin, _, _ = spin_at(db, user, 3, index_of(3, "5 Legacy Tokens"), wager=10)
    assert spin.prize_qty == 50
    assert main.award_totals(db, user.id)[user.id]["legacy_token"] == 50


def test_a_wager_feeds_the_pot(db):
    """A big wager fills the pot faster: the whole price of the spin goes in."""
    user = mkuser(db, "u", bits=10 ** 9, rank="Silver")
    user.test_unlocked = True
    db.commit()
    wager_ready(db, user, 3)
    main.start_spin(db, user, 10, wheel=3)
    assert main.jackpot_value(db) == 300_000
    main.start_spin(db, user, 1, wheel=3)
    assert main.jackpot_value(db) == 330_000


def test_the_jackpot_pays_the_pot_not_a_multiple_of_it(db):
    """The wager fills the pot, but it does not multiply the payout: a x10 takes the pot as
    it stands."""
    user = mkuser(db, "u", bits=10 ** 9, rank="Silver")
    user.test_unlocked = True
    db.commit()
    wager_ready(db, user, 3)
    main.meta_set(db, "jackpot", "500000")
    spin = spin_at(db, user, 3, index_of(3, "Progressive Legacy Jackpot"), wager=10)[0]
    # the pot as it stood, plus this spin's own 300,000, and nothing times ten
    assert spin.payout == 500_000 + 300_000
    assert main.jackpot_value(db) == 0


def test_claiming_twice_pays_once(db):
    user = mkuser(db, "u")
    spin, cost, _ = main.start_spin(db, user, 1, wheel=1)
    later = main.utcnow() + dt.timedelta(seconds=main.REVEAL_AFTER + 1)
    main.reveal_spin(db, user, spin.id, now=later)
    first, fresh = main.claim_spin(db, user, spin.id, now=later)
    second, again = main.claim_spin(db, user, spin.id, now=later)
    assert fresh is True and again is False
    assert main.delta_for(db, user.id) == -cost + first.payout


def test_balance_reflects_the_ledger(db):
    user = mkuser(db, "u", bits=10_000)
    before = main.available_bits(db, user)
    spin, cost, balance = spin_at(db, user, 1, index_of(1, "1000 BITS"))
    assert balance == before - cost
    assert main.available_bits(db, user) == before - cost + 1000


def test_a_spin_left_before_it_stops_is_forfeited(db):
    user = mkuser(db, "u")
    spin, cost, _ = main.start_spin(db, user, 1, wheel=1)
    later = main.utcnow() + dt.timedelta(seconds=main.SPIN_TTL + 5)
    main.sweep(db, now=later)
    db.refresh(spin)
    assert spin.status == "forfeited"
    assert main.available_bits(db, user) == 10_000_000 - cost
    with pytest.raises(main.Refused):
        main.claim_spin(db, user, spin.id, now=later)


def test_the_reveal_cannot_be_asked_for_early(db):
    user = mkuser(db, "u")
    spin, cost, _ = main.start_spin(db, user, 1, wheel=1)
    with pytest.raises(main.Refused) as exc:
        main.reveal_spin(db, user, spin.id)
    assert "still turning" in exc.value.message


# ===================================================================== free spins
def test_wheel_two_hands_over_five_free_spins(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin, cost, _ = spin_at(db, user, 2, index_of(2, "5 More Spins"))
    assert spin.prize_qty == 5
    state = main.free_state(db, user, 2)
    assert state["count"] == 5 and state["wager"] == 1 and state["wheel"] == 2


def test_free_spins_belong_to_the_wheel_that_paid_them(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin_at(db, user, 2, index_of(2, "5 More Spins"))
    assert main.free_state(db, user, 1) is None
    assert main.free_state(db, user, 2)["count"] == 5
    assert main.free_state(db, user, 3) is None
    with pytest.raises(main.Refused):
        main.start_spin(db, user, 1, free=True, wheel=1)      # nothing on wheel one


def test_a_free_spin_costs_nothing_and_keeps_its_wager(db):
    user = mkuser(db, "u", bits=10 ** 9, rank="Administrator")
    user.test_unlocked = True
    db.commit()
    wager_ready(db, user, 2)
    spin_at(db, user, 2, index_of(2, "5 More Spins"), wager=4)
    before = main.available_bits(db, user)
    spin, cost, balance = main.start_spin(db, user, 999, free=True, wheel=2)
    assert cost == 0 and balance == before
    assert spin.wager == 4 and spin.cost == 0
    out, _ = finish(db, user, spin, index=index_of(2, "600 BITS"))
    assert out.payout == 600 * 4
    assert main.available_bits(db, user) == before + 2400


def test_wheel_one_free_spins_grow_to_ten(db):
    """Two at x1, four at x2, and ten from x5 upwards."""
    user = mkuser(db, "u", bits=10 ** 9)
    wager_ready(db, user, 1)
    spin, _, _ = spin_at(db, user, 1, index_of(1, "2 More Spins"), wager=2)
    assert spin.prize_qty == 4
    spin, _, _ = spin_at(db, user, 1, index_of(1, "2 More Spins"), wager=5)
    assert spin.prize_qty == 10
    spin, _, _ = spin_at(db, user, 1, index_of(1, "2 More Spins"), wager=1000)
    assert spin.prize_qty == 10                     # the cap, not two thousand


def test_a_free_spin_can_win_more_free_spins(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin_at(db, user, 2, index_of(2, "5 More Spins"))
    spin, _, _ = main.start_spin(db, user, 1, free=True, wheel=2)
    finish(db, user, spin, index=index_of(2, "5 More Spins"))
    assert main.free_state(db, user, 2)["count"] == 5 + 5 - 1


def test_an_abandoned_free_spin_is_still_burned(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin_at(db, user, 2, index_of(2, "5 More Spins"))
    main.start_spin(db, user, 1, free=True, wheel=2)
    later = main.utcnow() + dt.timedelta(seconds=main.SPIN_TTL + 5)
    main.sweep(db, now=later)
    assert main.free_state(db, user, 2)["count"] == 4


def test_free_spins_stack_and_queue_oldest_first(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    main.grant_free_spins(db, user.id, 2, wager=1, count=5)
    main.grant_free_spins(db, user.id, 2, wager=3, count=2)
    state = main.free_state(db, user, 2)
    assert state["count"] == 5 and state["total"] == 7
    assert state["queue"] == [{"wager": 1, "count": 5}, {"wager": 3, "count": 2}]


def test_free_spins_of_two_wheels_do_not_share_a_queue(db):
    user = mkuser(db, "u", bits=10 ** 9)
    main.grant_free_spins(db, user.id, 1, wager=1, count=2)
    main.grant_free_spins(db, user.id, 2, wager=5, count=3)
    assert main.free_state(db, user, 1)["count"] == 2
    assert main.free_state(db, user, 2)["count"] == 3
    assert set(main.free_state_all(db, user)) == {"1", "2"}


def test_free_spins_are_never_an_obligation(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin_at(db, user, 2, index_of(2, "5 More Spins"))
    assert main.prize_rows(db) == []


# ===================================================================== the jackpot
def test_every_paid_wheel_three_spin_feeds_the_pot(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    assert main.jackpot_value(db) == 0
    spin_at(db, user, 3, index_of(3, "Bust - 0 BITS"))
    assert main.jackpot_value(db) == 30_000
    spin_at(db, user, 3, index_of(3, "1,500 BITS"))
    assert main.jackpot_value(db) == 60_000


def test_a_free_spin_does_not_feed_the_pot(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    main.grant_free_spins(db, user.id, 3, wager=1, count=1)
    spin, cost, _ = main.start_spin(db, user, 1, free=True, wheel=3)
    assert cost == 0
    assert main.jackpot_value(db) == 0


def test_a_forfeit_keeps_its_money_in_the_pot(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    main.start_spin(db, user, 1, wheel=3)
    later = main.utcnow() + dt.timedelta(seconds=main.SPIN_TTL + 5)
    main.sweep(db, now=later)
    assert main.jackpot_value(db) == 30_000


def test_landing_on_the_jackpot_takes_the_whole_pot_and_resets_it(db):
    one = mkuser(db, "one", bits=10 ** 9)
    two = mkuser(db, "two", bits=10 ** 9)
    for u in (one, two):
        u.test_unlocked = True
    db.commit()
    spin_at(db, one, 3, index_of(3, "Bust - 0 BITS"))
    spin_at(db, two, 3, index_of(3, "Bust - 0 BITS"))
    assert main.jackpot_value(db) == 60_000

    before = main.available_bits(db, one)
    spin, cost, _ = spin_at(db, one, 3, index_of(3, "Progressive Legacy Jackpot"))
    assert spin.payout == 90_000                # both earlier spins plus this one's price
    assert main.jackpot_value(db) == 0
    assert main.available_bits(db, one) == before - 30_000 + 90_000


def test_the_jackpot_cannot_be_collected_twice(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin, cost, _ = main.start_spin(db, user, 1, wheel=3)
    force_index(db, spin, index_of(3, "Progressive Legacy Jackpot"))
    later = main.utcnow() + dt.timedelta(seconds=main.REVEAL_AFTER + 1)
    main.reveal_spin(db, user, spin.id, now=later)
    first, _ = main.claim_spin(db, user, spin.id, now=later)
    second, fresh = main.claim_spin(db, user, spin.id, now=later)
    assert fresh is False and second.payout == first.payout == 30_000
    assert main.jackpot_value(db) == 0


def test_a_jackpot_that_was_never_collected_leaves_the_pot_alone(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin, cost, _ = main.start_spin(db, user, 1, wheel=3)
    force_index(db, spin, index_of(3, "Progressive Legacy Jackpot"))
    later = main.utcnow() + dt.timedelta(seconds=main.SPIN_TTL + 5)
    main.sweep(db, now=later)
    assert main.jackpot_value(db) == 30_000


# ===================================================================== fragments
def test_one_fragment_is_not_a_booster_yet(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin, _, _ = spin_at(db, user, 2, index_of(2, "Fragment of x5 BITS Booster"))
    assert spin.note == "1/3"
    assert main.fragments_of(db, user.id) == {"frag_x5": 1}
    assert main.award_totals(db, user.id) == {}


def test_three_fragments_become_one_booster(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    for _ in range(3):
        spin, _, _ = spin_at(db, user, 2, index_of(2, "Fragment of x2 BITS Booster"))
    assert spin.note == "+1 booster"
    assert main.fragments_of(db, user.id) == {}
    assert main.award_totals(db, user.id)[user.id]["bits_x2"] == 1


def test_fragments_keep_the_change_when_two_boosters_come_out(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    for _ in range(7):
        spin_at(db, user, 2, index_of(2, "Fragment of x3 BITS Booster"))
    assert main.fragments_of(db, user.id) == {"frag_x3": 1}
    assert main.award_totals(db, user.id)[user.id]["bits_x3"] == 2


def test_fragment_types_do_not_mix(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    for label in ("Fragment of x2 BITS Booster", "Fragment of x2 BITS Booster", "Fragment of x10 BITS Booster"):
        spin_at(db, user, 2, index_of(2, label))
    assert main.fragments_of(db, user.id) == {"frag_x2": 2, "frag_x10": 1}
    assert main.award_totals(db, user.id) == {}


def test_the_booster_made_from_fragments_is_a_real_obligation(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    for _ in range(3):
        spin_at(db, user, 2, index_of(2, "Fragment of x10 BITS Booster"))
    rows = main.prize_rows(db)
    assert [(r["key"], r["qty"]) for r in rows] == [("bits_x10", 1)]
    assert rows[0]["name"] == "x10 BITS Booster"


# ===================================================================== obligations
def test_mark_applied_clears_assets_but_never_fragments(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    for _ in range(3):
        spin_at(db, user, 2, index_of(2, "Fragment of x2 BITS Booster"))     # 1 booster, 0 left
    spin_at(db, user, 2, index_of(2, "Fragment of x5 BITS Booster"))         # 1 fragment left
    assert main.fragments_of(db, user.id) == {"frag_x5": 1}
    assert main.prize_rows(db)

    main.meta_set(db, "prizes_applied_at", main.utcnow().isoformat())
    db.commit()
    assert main.prize_rows(db) == []
    assert main.fragments_of(db, user.id) == {"frag_x5": 1}


def test_an_asset_won_after_mark_applied_shows_again(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    # pretend everything was handed out a second ago
    main.meta_set(db, "prizes_applied_at",
                  (main.utcnow() - dt.timedelta(seconds=1)).isoformat())
    db.commit()
    spin_at(db, user, 3, index_of(3, "x3 Legacy Booster"))
    assert [r["key"] for r in main.prize_rows(db)] == ["legacy_x3"]
    # and one more Mark applied takes it away again
    main.meta_set(db, "prizes_applied_at", (main.utcnow() + dt.timedelta(seconds=1)).isoformat())
    db.commit()
    assert main.prize_rows(db) == []


def test_obligations_combine_all_three_wheels(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin_at(db, user, 1, index_of(1, "Moving Logo"))
    spin_at(db, user, 3, index_of(3, "x5 Legacy Booster"))
    spin_at(db, user, 3, index_of(3, "1 Legacy Token"))
    keys = [r["key"] for r in main.prize_rows(db)]
    assert keys == ["moving_logo", "legacy_x5", "legacy_token"]


def test_legacy_tokens_are_owed_as_a_number(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin_at(db, user, 3, index_of(3, "5 Legacy Tokens"))
    assert main.prize_rows(db)[0]["qty"] == 5


def test_items_won_before_this_version_are_still_owed(db):
    """A database from the single-wheel site has no awards table. Nothing may go missing."""
    user = mkuser(db, "u")
    db.add(main.Spin(id="old-spin", user_id=user.id, wheel=1, wager=2, cost=600,
                     status="settled", kind="prize", label="Moving Logo", prize_qty=2,
                     created_at=main.utcnow(), settled_at=main.utcnow()))
    db.commit()
    assert main.backfill_awards(db) == 1
    assert main.backfill_awards(db) == 0                       # and it only happens once
    rows = main.prize_rows(db)
    assert [(r["key"], r["qty"]) for r in rows] == [("moving_logo", 2)]


# ===================================================================== admin rows
def test_the_players_table_drops_staked(db):
    assert [c for c, _ in main.PLAYER_COLS] == [
        "Player", "Rank", "Bank", "Delta", "Available", "Spins", "Last Spin", "Account"]
    assert "Staked" not in [c for c, _ in main.PLAYER_COLS]


def test_the_last_spin_column_knows_which_wheel(db):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin_at(db, user, 1, index_of(1, "500 BITS"))
    spin_at(db, user, 3, index_of(3, "90,000 BITS"))
    row = main.admin_rows(db)[0]
    assert row["last_wheel"] == 3
    assert row["last_label"] == "90,000 BITS"


def test_admin_rows_carry_the_wheels_the_player_has(db):
    user = mkuser(db, "u", bits=10 ** 9)
    _spin_n(db, user, 1, 100)
    row = main.admin_rows(db)[0]
    assert row["wheels"] == [1, 2]
    assert row["counts"] == {1: 100, 2: 0, 3: 0}


def test_admin_rows_carry_a_sortable_rank(db):
    mkuser(db, "a", rank="Bronze")
    mkuser(db, "b", rank="Organiser")
    rows = {r["display"]: r for r in main.admin_rows(db)}
    assert rows["a"]["rank_ord"] > rows["b"]["rank_ord"]


def test_prize_obligations_are_aggregated_per_player_and_item(db):
    one = mkuser(db, "one", bits=10 ** 9)
    two = mkuser(db, "two", bits=10 ** 9)
    spin_at(db, one, 1, index_of(1, "Static Logo"))
    spin_at(db, one, 1, index_of(1, "Static Logo"))
    spin_at(db, two, 1, index_of(1, "Static Logo"))
    rows = {(r["display"], r["key"]): r["qty"] for r in main.prize_rows(db)}
    assert rows == {("one", "static_logo"): 2, ("two", "static_logo"): 1}


# ===================================================================== the pages
def test_healthz(client):
    r = client.get("/healthz")
    assert r.status_code == 200 and r.json()["ok"] is True


def test_anonymous_gets_login_page(client):
    r = client.get("/")
    assert r.status_code == 200 and "Continue with Lichess" in r.text


def test_the_player_page_shows_three_tabs(db, client):
    user = mkuser(db, "u")
    login(client, user.id)
    r = client.get("/")
    assert r.status_code == 200
    html = r.text
    assert "id=tabs" in html
    assert "id=progBar" in html
    assert "How this wheel works" in html
    assert "Progressive Legacy Jackpot" in html
    assert '"numeral":"II"' in html or '"numeral": "II"' in html


def test_the_page_carries_the_odds_so_nobody_has_to_guess(db, client):
    user = mkuser(db, "u")
    login(client, user.id)
    html = client.get("/").text
    assert '"pct":45' in html or '"pct":45.0' in html        # Wheel II bust
    assert '"pct":0.05' in html                              # the jackpot
    assert "x20 BITS Booster" in html


def test_the_locked_tab_carries_the_count(db, client):
    user = mkuser(db, "u", bits=10 ** 9)
    _spin_n(db, user, 1, 36)
    login(client, user.id)
    html = client.get("/").text
    assert '"have":36' in html or '"have": 36' in html
    assert '"need":100' in html or '"need": 100' in html


def test_the_third_bar_says_it_is_secret(db, client):
    user = mkuser(db, "u", bits=10 ** 9)
    login(client, user.id)
    html = client.get("/").text
    assert '"secret":true' in html or '"secret": true' in html
    assert "not written down anywhere" in html


def test_the_bits_sentence_is_plain(db):
    user = mkuser(db, "u")
    payload = main.player_payload(db, user)
    assert "unchanged" in main.delta_line(0)
    assert "plus" in main.delta_line(500)
    assert "minus" in main.delta_line(-500)


def test_the_model_knows_every_column_the_migration_adds(db):
    """A column added by ALTER but missing from the model is written nowhere and read as
    None. That is how the last one slipped through the first time."""
    for table, column, _kind in main.COLUMN_ADDITIONS:
        assert column in main.Base.metadata.tables[table].columns, (table, column)


def test_mark_applied_still_clears_the_deltas(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    user = mkuser(db, "u", bits=10 ** 9)
    db.add(main.Ledger(user_id=user.id, delta=500, reason="test"))
    db.commit()
    login(client, admin.id)
    client.post("/admin/apply")
    db.expire_all()
    row = db.get(main.User, user.id)
    assert row.applied_through >= 1               # the ledger is behind us now
    assert main.player_payload(db, user)["delta"] == 0


def test_one_bad_column_does_not_stop_the_others(db, monkeypatch):
    """The deploy that broke: one rejected ALTER used to abandon the whole migration."""
    with main.engine.begin() as conn:
        conn.execute(main.text("ALTER TABLE users DROP COLUMN test_unlocked"))
        conn.execute(main.text("ALTER TABLE spins DROP COLUMN wheel"))
    good = main.column_ddl

    def fake(table, column, kind, dialect="sqlite"):
        if column == "test_unlocked":
            return "ALTER TABLE users ADD COLUMN"        # a syntax error, not a type
        return good(table, column, kind, dialect)

    monkeypatch.setattr(main, "column_ddl", fake)
    assert main.ensure_columns() is False         # it says so, rather than pretending
    assert "wheel" in {c["name"] for c in main.inspect(main.engine).get_columns("spins")}

    monkeypatch.setattr(main, "column_ddl", good)
    assert main.ensure_columns() is True
    assert "test_unlocked" in {c["name"] for c in main.inspect(main.engine).get_columns("users")}


def test_the_local_login_answers_with_a_see_other(db, client, monkeypatch):
    """307 would make the browser re-POST the form to "/" and land on a 405."""
    monkeypatch.setattr(main, "ALLOW_LOCAL", True)
    r = client.post("/auth/local", data={"username": "newbie", "password": "password123",
                                         "mode": "register"}, follow_redirects=False)
    assert r.status_code == 303 and r.headers["location"] == "/"


def test_admin_page_blocks_players(db, client):
    user = mkuser(db, "u")
    login(client, user.id)
    r = client.get("/admin")
    assert r.status_code == 403


def test_admin_sees_the_players_table(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    mkuser(db, "someone")
    login(client, admin.id)
    html = client.get("/admin").text
    for head in ("Player", "Rank", "Bank", "Delta", "Available", "Spins", "Last Spin", "Account"):
        assert head in html
    assert ">Staked<" not in html and "Staked<span" not in html


def test_admin_players_headings_are_sortable(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    login(client, admin.id)
    html = client.get("/admin").text
    assert html.count('class=sortable data-col=') >= len(main.PLAYER_COLS) + len(main.PRIZE_COLS)


def test_the_obligations_table_is_sortable_too(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    user = mkuser(db, "u", bits=10 ** 9)
    spin_at(db, user, 1, index_of(1, "Static Logo"))
    login(client, admin.id)
    html = client.get("/admin").text
    assert "id=prizeTbl" in html
    block = html.split("id=prizeTbl")[1].split("</table>")[0]
    assert block.count("data-col=") == 3
    assert 'data-v="static logo"' in html


def test_the_rank_column_sorts_by_the_ladder(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    bronze = mkuser(db, "bronze", rank="Bronze")
    organ = mkuser(db, "organ", rank="Organiser")
    login(client, admin.id)
    html = client.get("/admin").text
    assert ('data-v="%d"' % main.rank_ordinal("Bronze")) in html
    assert ('data-v="%d"' % main.rank_ordinal("Organiser")) in html
    assert main.rank_ordinal("Bronze") > main.rank_ordinal("Organiser")


def test_the_last_spin_cell_shows_a_wheel_number(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    user = mkuser(db, "u", bits=10 ** 9)
    spin_at(db, user, 1, index_of(1, "500 BITS"))
    login(client, admin.id)
    html = client.get("/admin").text
    assert "wtag w1" in html


def test_unlock_all_wheels_is_a_switch(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    login(client, admin.id)
    assert admin.test_unlocked is False
    client.post("/admin/unlock")
    db.refresh(admin)
    assert admin.test_unlocked is True
    client.post("/admin/unlock")
    db.refresh(admin)
    assert admin.test_unlocked is False


def test_the_jackpot_can_be_reset_from_the_admin_page(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    main.meta_set(db, "jackpot", "123456")
    db.commit()
    login(client, admin.id)
    client.post("/admin/jackpot")
    db.expire_all()
    assert main.jackpot_value(db) == 0


def test_mark_applied_marks_the_ledger(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    user = mkuser(db, "u", bits=10 ** 9)
    spin_at(db, user, 1, index_of(1, "500 BITS"))
    login(client, admin.id)
    assert "500" in client.get("/admin").text
    client.post("/admin/apply")
    db.refresh(user)
    assert user.applied_through > 0
    assert main.delta_for(db, user.id, user.applied_through) == 0


def test_admin_upload_workbook(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    user = mkuser(db, "known", rank="Herald", bits=0)
    login(client, admin.id)
    blob = workbook([(3, "known", 12_000_000, 0, "Sentinel", 0)])
    r = client.post("/admin/upload", files={"file": ("b.xlsx", blob, "application/octet-stream")})
    assert r.status_code in (200, 303)
    db.refresh(user)
    assert user.bank_bits == 12_000_000 and user.bank_rank == "Sentinel"


# ===================================================================== admin popup
def test_the_popup_knows_the_wheel_of_every_spin(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin_at(db, user, 1, index_of(1, "500 BITS"))
    spin_at(db, user, 3, index_of(3, "90,000 BITS"))
    login(client, admin.id)
    d = client.get("/admin/player/%d.json" % user.id).json()
    assert [s["wheel"] for s in d["spins"]] == [3, 1]


def test_the_popup_shows_fragments_assets_and_progress(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin_at(db, user, 2, index_of(2, "Fragment of x2 BITS Booster"))
    spin_at(db, user, 3, index_of(3, "x3 Legacy Booster"))
    login(client, admin.id)
    d = client.get("/admin/player/%d.json" % user.id).json()
    assert d["fragments"] == [{"key": "frag_x2", "name": "x2 BITS Booster fragment",
                               "emoji": main.ITEMS["bits_x2"]["emoji"], "count": 1}]
    assert [o["key"] for o in d["owed"]] == ["legacy_x3"]
    assert [o["key"] for o in d["ever"]] == ["legacy_x3"]
    assert [w["id"] for w in d["progress"]] == [1, 2, 3]
    assert d["progress"][0]["need"] == 50


def test_the_popup_is_admin_only(db, client):
    user = mkuser(db, "u")
    login(client, user.id)
    assert client.get("/admin/player/%d.json" % user.id).status_code == 403


def test_live_json_is_admin_only(db, client):
    user = mkuser(db, "u")
    login(client, user.id)
    assert client.get("/admin/live.json").status_code == 403


def test_live_json_reports_who_is_here(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    other = mkuser(db, "other")
    login(client, admin.id)
    client.get("/api/ping")
    d = client.get("/admin/live.json").json()
    assert d["players"][str(admin.id)]["active"] is True
    assert d["players"][str(other.id)]["active"] is False


def test_ping_stamps_last_seen(db, client):
    user = mkuser(db, "u")
    login(client, user.id)
    assert client.get("/api/ping").json()["ok"] is True
    db.refresh(user)
    assert user.last_seen is not None


# ===================================================================== the player api
def test_api_me_needs_a_login(client):
    assert client.get("/api/me").status_code == 401


def test_api_me_returns_all_three_wheels(db, client):
    user = mkuser(db, "u")
    login(client, user.id)
    d = client.get("/api/me").json()
    assert [w["id"] for w in d["wheels"]] == [1, 2, 3]
    assert d["wheels"][0]["cost"] == 300 and d["wheels"][2]["cost"] == 30000
    assert d["multi"] is False and d["jackpot"] == 0      # nothing open until the spins


def test_api_me_carries_the_gate_of_each_wheel(db, client):
    user = mkuser(db, "u", rank="Silver")
    login(client, user.id)
    d = client.get("/api/me").json()
    assert [(w["wagerRank"], w["wagerAfter"]) for w in d["wheels"]] == [
        ("Sovereign", 10), ("Administrator", 30), ("Silver", 150)]
    assert [w["multi"] for w in d["wheels"]] == [False, False, False]   # no spins yet
    wager_ready(db, user, 1)
    d = client.get("/api/me").json()
    assert d["wheels"][0]["multi"] is True and d["multi"] is True
    assert d["wheels"][1]["multi"] is False


def test_api_me_carries_the_bar_of_each_wheel(db, client):
    user = mkuser(db, "u")
    login(client, user.id)
    d = client.get("/api/me").json()
    third = d["wheels"][2]["progress"]
    assert third["secret"] is True and third["unit"] == "spins"
    assert d["wheels"][0]["progress"]["opens"] == "Wheel II"


def test_api_spin_and_claim_walk_the_whole_way(db, client):
    user = mkuser(db, "u", bits=10_000)
    wager_ready(db, user, 1)
    login(client, user.id)
    r = client.post("/api/spin", json={"wheel": 1, "wager": 2})
    assert r.status_code == 200
    j = r.json()
    assert j["cost"] == 600 and j["wheel"] == 1
    spin = db.get(main.Spin, j["spin_id"])
    seed_for(1, index_of(1, "1000 BITS"))
    spin.seed = seed_for(1, index_of(1, "1000 BITS"))
    spin.commit = main.commit_of(spin.seed)
    db.commit()
    import time
    time.sleep(0.05)
    later = main.utcnow() + dt.timedelta(seconds=main.REVEAL_AFTER + 1)
    main.reveal_spin(db, user, spin.id, now=later)          # visible to the claim call
    d = client.post("/api/spin/%s/claim" % spin.id).json()
    assert d["paid"] == 2000 and d["label"] == "1000 BITS"


def test_api_refuses_a_locked_wheel(db, client):
    user = mkuser(db, "u", bits=10 ** 9)
    login(client, user.id)
    r = client.post("/api/spin", json={"wheel": 2, "wager": 1})
    assert r.status_code == 400
    assert "locked" in r.json()["detail"]


def test_api_spins_lists_the_wheel_badge(db, client):
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin_at(db, user, 2, index_of(2, "600 BITS"))
    login(client, user.id)
    rows = client.get("/api/spins").json()
    assert rows[0]["wheel"] == 2 and ">II<" in rows[0]["badge"]


# ===================================================================== csv files
def test_deltas_csv_has_h_to_r_in_the_right_order(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    login(client, admin.id)
    text = client.get("/admin/deltas.csv").text
    head = text.splitlines()[0].split(",")
    assert head[:7] == ["username", "rank", "spending_before", "delta_to_type_into_H",
                        "spending_after", "spins", "staked"]
    assert head[7:10] == ["static_logos", "moving_logos", "custom_prefixes"]
    assert head[10:] == ["bits_x2_booster", "bits_x3_booster", "bits_x5_booster",
                         "bits_x10_booster", "bits_x20_booster", "legacy_x3_booster",
                         "legacy_x5_booster", "legacy_tokens"]
    assert len(head) == 18                                   # A to R


def test_deltas_csv_counts_the_assets(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin_at(db, user, 1, index_of(1, "Static Logo"))
    spin_at(db, user, 3, index_of(3, "x10 BITS Booster"))
    spin_at(db, user, 3, index_of(3, "3 Legacy Tokens"))
    login(client, admin.id)
    text = client.get("/admin/deltas.csv").text
    row = [l for l in text.splitlines() if l.startswith("@u,")][0].split(",")
    assert row[7:10] == ["1", "0", "0"]                      # logos and prefixes
    assert row[10:17] == ["0", "0", "0", "1", "0", "0", "0"]  # x2 x3 x5 x10 x20 lx3 lx5
    assert row[17] == "3"                                    # legacy tokens


def test_deltas_csv_leaves_fragments_out(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin_at(db, user, 2, index_of(2, "Fragment of x2 BITS Booster"))
    login(client, admin.id)
    text = client.get("/admin/deltas.csv").text
    row = [l for l in text.splitlines() if l.startswith("@u,")][0].split(",")
    assert row[10:18] == ["0"] * 8


def test_spins_csv_carries_the_wheel(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    user = mkuser(db, "u", bits=10 ** 9)
    user.test_unlocked = True
    db.commit()
    spin_at(db, user, 3, index_of(3, "90,000 BITS"))
    login(client, admin.id)
    text = client.get("/admin/spins.csv").text
    assert text.splitlines()[0].split(",")[2] == "wheel"
    assert ",III," in text


def test_prizes_csv_lists_the_items(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    user = mkuser(db, "u", bits=10 ** 9)
    spin_at(db, user, 3, index_of(3, "x5 Legacy Booster")) if False else None
    user.test_unlocked = True
    db.commit()
    spin_at(db, user, 1, index_of(1, "Custom Prefix"))
    login(client, admin.id)
    lines = client.get("/admin/prizes.csv").text.splitlines()
    assert lines[0] == "username,prize,owed"
    assert lines[1] == "@u,Custom Prefix,1"


def test_the_prize_csv_keeps_emoji_out(db, client):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    user = mkuser(db, "u", bits=10 ** 9)
    spin_at(db, user, 1, index_of(1, "Static Logo"))
    login(client, admin.id)
    text = client.get("/admin/prizes.csv").text
    assert "\U0001F5BC" not in text


# ===================================================================== upgrades
def test_ensure_columns_is_idempotent(db):
    assert main.ensure_columns() is True
    assert main.ensure_columns() is True
    names = {c["name"] for c in main.inspect(main.engine).get_columns("spins")}
    assert {"wheel", "note"} <= names


def test_ensure_columns_adds_every_column_it_promises(db):
    main.ensure_columns()
    for table, column, _kind in main.COLUMN_ADDITIONS:
        names = {c["name"] for c in main.inspect(main.engine).get_columns(table)}
        assert column in names, (table, column)


@pytest.mark.parametrize("dialect", ["sqlite", "postgresql"])
def test_a_boolean_default_is_spelled_the_same_everywhere(dialect):
    """PostgreSQL rejects BOOLEAN DEFAULT 0 outright. That one line broke the first deploy."""
    ddl = main.column_ddl("users", "test_unlocked", "bool", dialect)
    assert "BOOLEAN DEFAULT FALSE" in ddl
    assert "DEFAULT 0" not in ddl


@pytest.mark.parametrize("dialect,stamp", [
    ("sqlite", "TIMESTAMP"), ("postgresql", "TIMESTAMP WITH TIME ZONE")])
def test_the_timestamp_type_follows_the_database(dialect, stamp):
    ddl = main.column_ddl("users", "last_seen", "stamp", dialect)
    assert ddl == "ALTER TABLE users ADD COLUMN last_seen " + stamp


def test_every_addition_is_a_plain_add_column():
    for table, column, kind in main.COLUMN_ADDITIONS:
        assert main.column_ddl(table, column, kind, "postgresql").startswith(
            "ALTER TABLE %s ADD COLUMN %s " % (table, column))


def test_an_old_database_keeps_its_rows(db):
    """A database made before wheels existed gets wheel=1 and loses nothing."""
    user = mkuser(db, "old")
    db.add(main.Spin(id="s1", user_id=user.id, wager=1, cost=300, status="settled",
                     kind="bits", label="500 BITS", payout=500,
                     created_at=main.utcnow(), settled_at=main.utcnow()))
    db.commit()
    main.ensure_columns()
    row = db.get(main.Spin, "s1")
    assert int(row.wheel or 1) == 1
    assert main.wheel_counts(db, user.id)[1] == 1


def test_the_admin_meets_the_emoji_fallback(client):
    r = client.get("/emoji-sheet.png")
    assert r.status_code == 200 and r.content[:8] == b"\x89PNG\r\n\x1a\n"


# ===================================================================== the page itself
def test_the_chances_page_does_not_talk_the_wheel_down(db, client):
    """A note giving the expected return as a percentage of the price reads as a warning and
    stops people spinning. The odds themselves stay; that sentence goes."""
    user = mkuser(db, "u")
    login(client, user.id)
    html = client.get("/").text
    assert "oddsFoot" not in html and "oddsFoot" not in main.WHEEL_JS
    for phrase in ("of the price", "hands back about", "great many spins"):
        assert phrase not in main.WHEEL_JS


def test_every_section_of_the_explainer_has_one_heading():
    """Two h4s in a row happened once; nothing catches that by looking at the page."""
    import re as _re
    how = main.HOW
    for chunk in how.split("<section")[1:]:
        heads = _re.findall(r"<h4>", chunk)
        assert len(heads) <= 1, chunk[:120]
    assert how.count("<h4>What the bar counts</h4>") == 1


def test_the_chances_are_not_struck_through(db, client):
    """The rows fell back to bare <s> when the card was renamed and every percentage came out
    crossed out. <s> carries a strikethrough of its own, so the rule has to reach it."""
    css = main.CSS
    assert ".oddsbox li .pct{" in css                # a plain span, not <s>
    assert ".odds li" not in css                     # the old prefix matches nothing now
    assert ".oddsbox ol{list-style:none" in css      # and the list itself is still styled
    assert "'</b><span class=pct>'" in main.WHEEL_JS  # <s> carries a strikethrough of its own


def test_the_fragment_slices_say_what_they_become():
    labels = [s["label"] for s in main.WHEEL_2 if s["key"] and s["key"].startswith("frag_")]
    assert labels == ["Fragment of x2 BITS Booster", "Fragment of x3 BITS Booster",
                      "Fragment of x5 BITS Booster", "Fragment of x10 BITS Booster"]
    for s in main.WHEEL_2:
        if s["key"] and s["key"].startswith("frag_"):
            assert len(s["lines"]) == 2                 # still two short lines on the wheel


def test_the_explainer_keeps_later_wheels_to_themselves(db, client):
    """Wheel I must not mention the jackpot or the fragments, Wheel II must not mention the
    jackpot, and each section says which wheels it belongs to."""
    user = mkuser(db, "u")
    login(client, user.id)
    html = client.get("/").text
    how = html[html.index('id=howBox'):]
    how = how[:how.index("</details>")]
    jackpot = how[how.index("<h4>The Progressive Legacy Jackpot</h4>") - 40:]
    assert 'data-ow="3"' in jackpot[:40]                 # Wheel III only
    frags = how[how.index("<h4>Fragments</h4>") - 40:]
    assert 'data-ow="2"' in frags[:40]                   # Wheel II only
    boost = how[how.index("<h4>Boosters and tokens</h4>") - 40:]
    assert 'data-ow="3"' in boost[:40]                   # Wheel III only
    assert how.count("data-ow=") >= 8                    # every section is tagged
    assert "renderHow" in html                           # and the page filters on the tab


def test_the_wheel_one_section_never_names_a_fragment(db, client):
    user = mkuser(db, "u")
    login(client, user.id)
    html = client.get("/").text
    how = html[html.index('id=howBox'):]
    how = how[:how.index("</details>")]
    for chunk in how.split("<section"):
        head = chunk[:40]
        if 'data-ow="1"' in head:                        # the section shown on Wheel I
            body = chunk
            assert "jackpot" not in body.lower()
            assert "fragment" not in body.lower()


# ===================================================================== selling logos back
def test_the_exchange_rates_are_the_posted_ones():
    assert main.EXCHANGE == {"static_logo": 50, "moving_logo": 100, "custom_prefix": 150}
    assert main.EXCHANGE_ORDER == ["static_logo", "moving_logo", "custom_prefix"]


def test_the_exchange_lists_what_the_player_holds(db):
    user = mkuser(db, "u")
    main.give_award(db, user.id, "static_logo", 4)
    main.give_award(db, user.id, "custom_prefix", 2)
    db.commit()
    listed = {x["key"]: x["held"] for x in main.exchange_list(db, user)}
    assert listed == {"static_logo": 4, "moving_logo": 0, "custom_prefix": 2}
    assert [x["rate"] for x in main.exchange_list(db, user)] == [50, 100, 150]


def test_selling_logos_turns_them_into_bits(db):
    user = mkuser(db, "u", bits=0)
    main.give_award(db, user.id, "static_logo", 4)
    db.commit()
    bits = main.convert_assets(db, user, "static_logo", 3)
    assert bits == 150                       # three at fifty each
    assert main.assets_of(db, user.id)["static_logo"] == 1
    assert main.award_totals(db, user.id)[user.id]["static_logo"] == 1
    assert main.available_bits(db, user) == 150
    assert main.delta_for(db, user.id, 0) == 150


def test_selling_everything_empties_the_list(db):
    user = mkuser(db, "u")
    main.give_award(db, user.id, "moving_logo", 2)
    db.commit()
    assert main.convert_assets(db, user, "moving_logo", 2) == 200
    assert main.assets_of(db, user.id)["moving_logo"] == 0
    assert main.award_totals(db, user.id).get(user.id, {}) == {}
    assert main.exchange_list(db, user)[1]["held"] == 0


def test_the_exchange_takes_them_off_the_admin_list(db):
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    user = mkuser(db, "u")
    main.give_award(db, user.id, "custom_prefix", 1)
    db.commit()
    rows = [r for r in main.prize_rows(db) if r["key"] == "custom_prefix"]
    assert rows and rows[0]["qty"] == 1
    main.convert_assets(db, user, "custom_prefix", 1)
    rows = [r for r in main.prize_rows(db) if r["key"] == "custom_prefix"]
    assert rows == []                        # nothing left for the admin to hand over


def test_the_exchange_refuses_more_than_you_have(db):
    user = mkuser(db, "u")
    main.give_award(db, user.id, "static_logo", 1)
    db.commit()
    with pytest.raises(main.Refused) as exc:
        main.convert_assets(db, user, "static_logo", 2)
    assert "You have 1" in exc.value.message
    with pytest.raises(main.Refused):
        main.convert_assets(db, user, "static_logo", 0)
    with pytest.raises(main.Refused):
        main.convert_assets(db, user, "bits_x2", 1)          # not a logo
    assert main.assets_of(db, user.id)["static_logo"] == 1
    assert main.delta_for(db, user.id, 0) == 0


def test_nothing_can_be_sold_after_the_batch_is_applied(db, client):
    """The logos have been handed over by then, so the site no longer holds them."""
    admin = mkuser(db, CC, bits=10 ** 9, admin=True)
    user = mkuser(db, "u")
    main.give_award(db, user.id, "static_logo", 3)
    db.commit()
    login(client, admin.id)
    client.post("/admin/apply")
    db.expire_all()
    assert main.assets_of(db, user.id)["static_logo"] == 0
    with pytest.raises(main.Refused):
        main.convert_assets(db, user, "static_logo", 1)


def test_the_api_sells_logos(db, client):
    user = mkuser(db, "u", bits=0)
    main.give_award(db, user.id, "static_logo", 2)
    db.commit()
    login(client, user.id)
    d = client.post("/api/exchange", json={"key": "static_logo", "qty": 2}).json()
    assert d["bits"] == 100 and d["available"] == 100
    assert {x["key"]: x["held"] for x in d["exchange"]}["static_logo"] == 0
    bad = client.post("/api/exchange", json={"key": "static_logo", "qty": 1})
    assert bad.status_code == 400 and "You have 0" in bad.json()["detail"]


def test_the_page_carries_the_exchange_card(db, client):
    user = mkuser(db, "u")
    main.give_award(db, user.id, "static_logo", 5)
    db.commit()
    login(client, user.id)
    html = client.get("/").text
    assert "id=convCard" in html and "id=convRows" in html
    assert "\"rate\":50" in html or "\"rate\": 50" in html


# ===================================================================== the pot's picture
def test_the_jackpot_slice_has_its_own_picture():
    jackpot = [s for s in main.slices(3) if s["kind"] == "jackpot"][0]
    assert jackpot["img"] == "/jackpot.png"


def test_the_picture_is_served_as_a_png(client):
    r = client.get("/jackpot.png")
    assert r.status_code == 200
    assert r.headers["content-type"] == "image/png"
    assert len(r.content) > 4000
