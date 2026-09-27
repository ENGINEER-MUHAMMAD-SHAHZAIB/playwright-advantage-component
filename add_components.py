"""
Advantage SIS - add course components from Excel.

Reads data/automation_ready.xlsx (sheet "Components") and, for every row with Run = Yes,
does what your recordings did:
  open course (creates the COURSE and its GRADE BOOK first if they don't exist)
  -> course grid: pick component, number, weight, min score, sequence,
  Required -> row button -> Save -> Confirm
If the component does not exist yet in the list and CREATE_IF_MISSING = True, it first
creates it with New Component (Description / Type / Status / Campus -> Save).

SAFE DEFAULTS: LIMIT = 1 and DRY_RUN = True (fills the grid row but never clicks Save).

Run:   python add_components.py
Stop:  Ctrl + C in the terminal
"""
import csv
import re
import datetime
import glob
import os
import sys
import time
from collections import Counter

import pandas as pd
from dotenv import load_dotenv
from playwright.sync_api import TimeoutError as PWTimeout
from playwright.sync_api import sync_playwright

# ============================ SETTINGS (edit these) ============================
EXCEL_FILE = "data/automation_ready.xlsx"
SHEET = "Components"

LIMIT = 400                # max NEW rows this run (already-done rows are not counted)
DRY_RUN = False            # True = fill the grid row for 1 row, do NOT click Save
START_FROM_ROW_ID = 1       # normally 1 - the done-cache resumes automatically
ALLOW_LIVE = True        # must be True to SAVE on the live site advantage.fameinc.com

BROWSER_CHANNEL = "msedge"  # "msedge" (as in your recording) or "" for Playwright Chromium
SLOW_MO = 30                # pause between clicks in ms (0 = fastest, 250 = easy to watch)
HEADLESS = False            # True = hide the browser (a bit faster; you will not see it)

PROGRAM_SEARCH_TEXT = "Cosmetology"
PROGRAM_OPTION = "#programAndProgramVersions-option-3"   # the version you picked when recording

BATCH_SIZE = 10            # rows to add before ONE Save (1 = save after every row, like before)
PRECREATE_MISSING = True   # create all missing components for a course first, then add rows in batches
SKIP_DONE_FROM_LOGS = True  # skip rows already saved (done-cache file + old logs) - no searching again
VERIFY_ONLY = False        # True = only CHECK Advantage against the Excel (adds nothing), writes
                           #        logs/verify_*.csv and rebuilds the done-cache from what really exists
REDO_ROW_IDS = []          # Row IDs to add AGAIN even if the cache says done (delete them in Advantage first!)
                           # e.g. [340, 341, 342, 343, 344]  - set back to [] afterwards
MAX_RETRIES = 3            # if the page gets stuck: reload, reopen the course and try again (per row)
GRID_WAIT_MAX_S = 180      # max seconds to wait for a big course grid to load
GRID_STALL_S = 30          # give up waiting if nothing more loads for this many seconds
CREATE_COURSE_IF_MISSING = True   # course not in the ADD/VIEW COURSE list -> ADD NEW course + grade book
CREATE_IF_MISSING = True  # True = create the component with "New Component" if it is not in the list yet
GRID_ROW = 0               # grid row used for a new entry (both recordings used [0])
# The small "add row" button in the grid entry row (a button placed directly in a table cell).
# Its jss number changes between sessions (jss1280, jss1765...), so the script finds it by position.
ROW_BUTTON_SELECTOR = "td.MuiTableCell-body > button.MuiButtonBase-root"
# ===============================================================================

load_dotenv()
URL = os.getenv("SIS_URL", "https://advantage.fameinc.com/login")
USER = os.getenv("SIS_USER", "")
PASS = os.getenv("SIS_PASS", "")
TENANT = os.getenv("SIS_TENANT", "royal")
TENANT_OPTION = os.getenv("SIS_TENANT_OPTION", "RoyalTreatmentBeautyAcademy")

IS_LIVE = "advantage.fameinc.com" in URL
os.makedirs("logs", exist_ok=True)
STAMP = f"{datetime.datetime.now():%Y%m%d_%H%M%S}"
LOG_FILE = f"logs/components_{STAMP}.csv"


def say(msg):
    print(f"{datetime.datetime.now():%H:%M:%S}  {msg}", flush=True)


# ------------------------------------------------------------------ helpers
def pick_option(page, input_locator, value):
    """Open an autocomplete/dropdown, type the value, and click the exact option."""
    input_locator.click()
    try:
        input_locator.fill(value)
    except Exception:
        pass  # some dropdowns are read-only; the option list is enough
    option = page.get_by_role("option", name=value, exact=True)
    option.first.wait_for(timeout=8000)
    option.first.click()


def set_checkbox(box, want_yes):
    if box.is_checked() != want_yes:
        box.click()


# ------------------------------------------------------------------ steps (from your recording)
def login(page):
    say(f"Opening {URL}")
    for attempt in range(1, 4):          # the site is sometimes slow: try up to 3 times
        try:
            page.goto(URL, timeout=60000, wait_until="domcontentloaded")
            page.get_by_role("button", name="Username & password").wait_for(timeout=60000)
            break
        except PWTimeout:
            if attempt == 3:
                raise RuntimeError("Advantage login page did not load after 3 tries - check your internet/VPN and try again")
            say(f"Login page slow, trying again ({attempt}/3)...")
    page.get_by_role("button", name="Username & password").click()
    page.get_by_role("textbox", name="Username").fill(USER)
    page.get_by_role("textbox", name="Password").fill(PASS)
    page.get_by_role("button", name="Login").click()
    page.get_by_role("textbox", name="Tenant").fill(TENANT)
    page.get_by_role("option", name=TENANT_OPTION).click()
    page.get_by_role("button", name="Continue").click()
    say("Logged in")


def open_program(page):
    page.get_by_role("button", name="add").click()
    page.get_by_role("button", name="Setup").click()
    page.get_by_role("button", name="Academic", exact=True).click()
    page.get_by_role("button", name="Academics").nth(1).click()
    page.get_by_role("button", name="Programs").click()
    page.get_by_role("textbox", name="Search and Add Program and").click()
    page.locator(PROGRAM_OPTION).click()
    page.get_by_role("tab", name="Definition").click()
    say(f"Opened program: {PROGRAM_SEARCH_TEXT}")


def save_in_section(page, heading):
    """Click the Save button that belongs to the card with this heading
    (e.g. 'Course Details' or 'Grade Book Details')."""
    sec = page.locator("div").filter(has=page.get_by_text(heading, exact=True)) \
        .filter(has=page.get_by_role("button", name="Save"))
    sec.last.get_by_role("button", name="Save").first.click()
    confirm = page.get_by_role("button", name="Confirm")
    try:
        confirm.first.wait_for(timeout=2500)
        confirm.first.click()
    except PWTimeout:
        pass
    page.wait_for_timeout(800)


def fill_text(box, value):
    box.click()
    box.press("ControlOrMeta+a")
    box.fill(str(value))


class DryRunStop(Exception):
    pass


def create_course(page, row, log):
    """ADD NEW (course) -> Course *, Status *, Campus *, Minimum Score -> Save  (your recording)."""
    say(f"NEW COURSE  {row['course_list']}")
    page.get_by_role("button", name="ADD NEW").first.click()
    fill_text(page.get_by_role("textbox", name="Course *"), row["course_list"])
    dialog = page.get_by_role("dialog", name="Courses close")
    pick_option(page, dialog.locator("#statuses").first, row["status"])
    camp = dialog.locator("#campusId")
    if camp.count() and camp.first.input_value().strip() != row["campus"]:
        pick_option(page, camp.first, row["campus"])
    if row["course_min"] not in (None, ""):
        fill_text(page.get_by_role("textbox", name="Minimum Score"), row["course_min"])
    if DRY_RUN:
        page.screenshot(path="logs/dryrun_new_course.png")
        raise DryRunStop("Course Details filled, NOT saved (logs/dryrun_new_course.png)")
    save_in_section(page, "Course Details")
    page.get_by_role("cell", name=row["course_list"], exact=True).first.wait_for(timeout=20000)
    state_add("courses", row["course_list"])
    log.writerow([f"{datetime.datetime.now():%H:%M:%S}", row["id"], row["course"], row["course_list"], "COURSE CREATED", ""])
    say(f"OK    course created: {row['course_list']}")
    page.get_by_role("cell", name=row["course_list"], exact=True).first.click()


TAG_DATE_JS = """() => {
    document.querySelectorAll('[data-auto="effdate"]').forEach(e => e.removeAttribute('data-auto'));
    const labels = [...document.querySelectorAll('label')]
        .filter(l => /Effective start date/i.test(l.textContent) && l.offsetParent !== null);
    for (const l of labels) {
        let inp = l.htmlFor ? document.getElementById(l.htmlFor) : null;
        if (!inp) { const fc = l.closest('.MuiFormControl-root'); inp = fc && fc.querySelector('input'); }
        if (inp) { inp.setAttribute('data-auto', 'effdate'); return true; }
    }
    return false;
}"""


def date_input(page):
    """The visible 'Effective start date *' input of Grade Book Details."""
    if page.evaluate(TAG_DATE_JS):
        return page.locator('[data-auto="effdate"]').first
    raise RuntimeError("Effective start date field not found")


def set_date(page, value):
    """Your recording: click the date box and fill '01/01/2020'."""
    box = date_input(page)
    box.scroll_into_view_if_needed()
    box.click()
    box.fill(value)
    if value[-4:] not in box.input_value():          # masked input did not take fill(): type it
        box.press("ControlOrMeta+a")
        box.press("Backspace")
        box.press_sequentially(value.replace("/", ""), delay=40)
    got = box.input_value().replace("_", "").strip()
    if value[-4:] not in got:
        raise RuntimeError(f"Could not type the Effective start date ({value}), field shows '{got}'")
    return got


def create_grade_book(page, row, log):
    """Grade Book -> ADD NEW (if the form is not open yet) -> Grade book * (Subject - Program),
    Effective start date -> Save. If Advantage will not save an empty grade book, it is saved
    later together with the components (one Save)."""
    say(f"NEW GRADE BOOK  {row['course']}")
    name_box = page.get_by_role("textbox", name="Grade book *")
    if not name_box.count() or not name_box.first.is_visible():
        page.get_by_role("button", name="ADD NEW").nth(1).click()
    fill_text(page.get_by_role("textbox", name="Grade book *").first, row["course"])
    got = set_date(page, row["eff_date"])
    if DRY_RUN:
        page.screenshot(path="logs/dryrun_new_gradebook.png")
        raise DryRunStop(f"Grade Book Details filled (date {got}), NOT saved (logs/dryrun_new_gradebook.png)")
    page.get_by_role("button", name="Save").nth(1).click()        # your recording
    confirm = page.get_by_role("button", name="Confirm")
    try:
        confirm.first.wait_for(timeout=5000)
        confirm.first.click()
    except PWTimeout:
        pass
    page.wait_for_timeout(1000)
    state_add("gradebooks", row["course"])
    listed = page.get_by_role("cell", name=row["course"], exact=True)
    try:
        listed.first.wait_for(timeout=8000)
        log.writerow([f"{datetime.datetime.now():%H:%M:%S}", row["id"], row["course"], row["course"], "GRADE BOOK CREATED", got])
        say(f"OK    grade book created: {row['course']} (start {got})")
    except PWTimeout:
        log.writerow([f"{datetime.datetime.now():%H:%M:%S}", row["id"], row["course"], row["course"], "GRADE BOOK", "not saved empty - will be saved with the components"])
        say("NOTE  grade book not saved on its own - it will be saved together with the components")


def courses_dialog_open(page):
    d = page.get_by_role("dialog", name=re.compile(r"Courses", re.I))
    return d.count() > 0 and d.first.is_visible()


def go_to_course_list(page):
    """Show the Courses list. If the Courses window is already open (e.g. after working in
    another course) we stay in it - ADD/VIEW COURSE is behind it and cannot be clicked."""
    if not courses_dialog_open(page):
        page.get_by_role("button", name="ADD/VIEW COURSE").click()
    page.get_by_role("button", name="ADD NEW").first.wait_for(timeout=20000)
    page.wait_for_timeout(800)


def exact(text):
    return re.compile(r"^\s*" + re.escape(text.strip()) + r"\s*$", re.I)


def find_course_cell(page, row):
    """The course in the list, e.g. 'Practical Exams [Cosmetology]'. Uses the search box only
    if it is not visible straight away."""
    cell = page.get_by_role("cell", name=exact(row["course_list"]))
    if cell.count():
        return cell.first
    search = page.get_by_role("textbox", name="Search Courses")
    if not search.count():
        search = page.get_by_placeholder(re.compile(r"Search Courses", re.I))
    if search.count():
        fill_text(search.first, row["subject"])
        page.wait_for_timeout(1500)
        cell = page.get_by_role("cell", name=exact(row["course_list"]))
        if cell.count():
            return cell.first
    return None


def select_course(page, row, log, create=True):
    """Course from the Excel: exists -> click it (use it). Missing -> create it (if allowed).
    Then the same for its grade book ('Subject - Program')."""
    go_to_course_list(page)
    cell = find_course_cell(page, row)
    if cell is not None:
        cell.scroll_into_view_if_needed()
        cell.click()
        state_add("courses", row["course_list"])
        say(f"FOUND course: {row['course_list']} (selected)")
    elif create and CREATE_COURSE_IF_MISSING:
        create_course(page, row, log)
    else:
        raise RuntimeError(f"Course '{row['course_list']}' is not in the course list")

    gb = page.get_by_role("cell", name=exact(row["course"]))
    try:
        gb.first.wait_for(timeout=10000)
        gb.first.scroll_into_view_if_needed()
        gb.first.click()
        state_add("gradebooks", row["course"])
        say(f"FOUND grade book: {row['course']} (selected)")
    except PWTimeout:
        if create and CREATE_COURSE_IF_MISSING:
            create_grade_book(page, row, log)
        else:
            raise RuntimeError(f"Grade book '{row['course']}' is not in course '{row['course_list']}'")
    grid_inputs(page).first.wait_for(timeout=20000)
    say(f"Opened course: {row['course']}")


def open_course(page, row, log):
    select_course(page, row, log, create=True)


def grid_inputs(page):
    return page.locator('input[name^="gradeBookComponentTypeId["]')


GRID_JS = """els => els.map(e => [e.getAttribute('name') || '', (e.value || '').trim()])"""


def grid_values(page):
    """All grid rows in ONE browser call (fast even with 300+ rows): [(name_attr, value), ...]"""
    return grid_inputs(page).evaluate_all(GRID_JS)


SCROLL_JS = """async (els, newRow) => {
    for (const e of els) {
        if (!(e.value || '').trim() && e.getAttribute('name') !== newRow) {
            e.scrollIntoView({block: 'center'});
            await new Promise(r => setTimeout(r, 120));
        }
    }
    if (els.length) els[0].scrollIntoView({block: 'center'});
}"""


def wait_grid_loaded(page, strict=False):
    """Wait for the saved rows to show their names. Big grids load slowly and some rows only
    load when scrolled into view, so: wait while it makes progress, scroll the empty rows
    into view, and stop waiting when nothing more loads (the done-cache still prevents
    duplicates). Returns True when fully loaded."""
    new_row = f"gradeBookComponentTypeId[{GRID_ROW}]"
    start = time.time()
    last_count, stable = -1, 0
    best_empty, last_progress, last_scroll = None, time.time(), 0
    while True:
        vals = grid_values(page)
        empty = sum(1 for nm, v in vals if nm != new_row and v == "")
        spinning = page.locator(".MuiCircularProgress-root:visible").count()
        if vals and empty == 0 and spinning == 0 and len(vals) == last_count:
            stable += 1
            if stable >= 2:
                return True
        else:
            stable = 0
        if best_empty is None or empty < best_empty or len(vals) != last_count:
            best_empty, last_progress = empty, time.time()
        last_count = len(vals)
        now = time.time()
        if empty and now - last_progress > 4 and now - last_scroll > 8:
            grid_inputs(page).evaluate_all(SCROLL_JS, new_row)   # load rows that are out of view
            last_scroll = now
        if now - last_progress > GRID_STALL_S or now - start > GRID_WAIT_MAX_S:
            if strict:
                raise RuntimeError("Course grid did not finish loading")
            say(f"WARN  grid partly loaded ({len(vals) - empty} of {len(vals)} rows) - using the done-cache")
            return False
        page.wait_for_timeout(300)


def names_in_grid(page):
    """Counter of component names already saved in the course (lower-case)."""
    new_row = f"gradeBookComponentTypeId[{GRID_ROW}]"
    return Counter(v.lower() for nm, v in grid_values(page) if nm != new_row and v)


CACHE_FILE = os.path.join("data", "done_" + os.path.splitext(os.path.basename(EXCEL_FILE))[0] + ".csv")


VERIFIED_FLAG = CACHE_FILE.replace(".csv", ".verified")


def load_done():
    """(course, component) already saved. After a VERIFY run the cache holds exactly what exists
    in Advantage, so only the cache is used. Before that: cache + OK rows in old logs."""
    done = set()
    files = [CACHE_FILE]
    if not os.path.exists(VERIFIED_FLAG):
        files += [p for p in glob.glob("logs/components_*.csv")
                  if not p.replace("\\", "/").endswith(LOG_FILE.split("/")[-1])]
    for path in files:
        if not os.path.exists(path):
            continue
        try:
            with open(path, newline="", encoding="utf-8") as fh:
                for r in csv.DictReader(fh):
                    if path == CACHE_FILE or r.get("status") == "OK":
                        done.add((r["course"].strip().lower(), r["component"].strip().lower()))
        except Exception:
            pass
    return done


def cache_rewrite(keys_rows):
    """Replace the done-cache with rows that were really found in Advantage (VERIFY run)."""
    tmp = CACHE_FILE + ".tmp"
    with open(tmp, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["time", "row_id", "course", "component"])
        for r in keys_rows:
            w.writerow([f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S} verified", r["id"], r["course"], r["name"]])
    os.replace(tmp, CACHE_FILE)
    with open(VERIFIED_FLAG, "w") as fh:
        fh.write(f"verified {datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n")


def cache_add(rows):
    """Write saved rows to the done-cache immediately (safe even if the run stops next second)."""
    new = not os.path.exists(CACHE_FILE)
    with open(CACHE_FILE, "a", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        if new:
            w.writerow(["time", "row_id", "course", "component"])
        for r in rows:
            w.writerow([f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S}", r["id"], r["course"], r["name"]])
        fh.flush()
        os.fsync(fh.fileno())


STATE_FILE = os.path.join("data", "state_" + os.path.splitext(os.path.basename(EXCEL_FILE))[0] + ".json")


def load_state():
    """What already exists in Advantage (found or created by earlier runs):
    courses, grade books and components - so they are selected, not created or searched again."""
    import json
    try:
        with open(STATE_FILE, encoding="utf-8") as fh:
            data = json.load(fh)
        return {k: set(v) for k, v in data.items()}
    except Exception:
        return {"courses": set(), "gradebooks": set(), "components": set()}


def state_add(kind, name):
    import json
    STATE.setdefault(kind, set()).add(name.strip().lower())
    tmp = STATE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        json.dump({k: sorted(v) for k, v in STATE.items()}, fh, indent=1)
    os.replace(tmp, STATE_FILE)


STATE = load_state()


def empty_grid_index(page):
    """Grid row for the new entry: GRID_ROW if it is empty, otherwise the first empty row."""
    first = page.locator(f'input[name="gradeBookComponentTypeId[{GRID_ROW}]"]')
    if first.count() and first.input_value().strip() == "":
        return GRID_ROW
    boxes = grid_inputs(page)
    for i in range(boxes.count()):
        if boxes.nth(i).input_value().strip() == "":
            name = boxes.nth(i).get_attribute("name")        # e.g. gradeBookComponentTypeId[3]
            return name[name.index("[") + 1:name.index("]")]
    raise RuntimeError("No empty row found in the course grid")


def create_component(page, row):
    """New Component dialog: Description, System component type, Status, Campus, Save."""
    page.get_by_role("button", name="New Component").click()
    dialog = page.get_by_role("dialog", name="Courses close")
    dialog.locator("#description").fill(row["name"])
    pick_option(page, page.get_by_role("textbox", name="System component type *"), row["type"])
    form = page.locator("form").filter(has_text="Description *System component")
    pick_option(page, form.locator("#statuses"), row["status"])
    pick_option(page, form.locator("#campusId"), row["campus"])
    page.get_by_role("button", name="Save").nth(2).click()
    page.wait_for_timeout(400)


def select_component(page, n, name):
    """Type the name in the grid's component box and click the matching option. True if found."""
    comp = page.locator(f'input[name="gradeBookComponentTypeId[{n}]"]')
    comp.click()
    comp.fill(name)
    opt = page.get_by_role("option", name=name, exact=True)
    try:
        opt.first.wait_for(timeout=8000)
    except PWTimeout:
        comp.press("Escape")
        comp.fill("")
        return False
    opt.first.click()
    return True


def fill_grid_row(page, row, pending):
    """Fill the new-entry grid row (from your recording) and click the row button that adds it
    to the list. Nothing is saved to Advantage until save_grid() is called.
    'pending' = rows added but not saved yet (saved first if a New Component must be created)."""
    n = empty_grid_index(page)
    created = ""
    if not select_component(page, n, row["name"]):
        if not CREATE_IF_MISSING or DRY_RUN:
            raise RuntimeError("Component not in the list. Create it first, or set CREATE_IF_MISSING = True")
        if pending:
            raise NeedsSaveFirst()   # save the rows in the batch before opening New Component
        create_component(page, row)
        created = "created + "
        n = empty_grid_index(page)
        if not select_component(page, n, row["name"]):
            raise RuntimeError("Component was created but is still not in the grid list")

    for field, value in (("number", row["number"]), ("weight", row["weight"]),
                         ("minimumScore", row["min"]), ("sequence", row["seq"])):
        box = page.locator(f'input[name="{field}[{n}]"]')
        box.fill("" if value is None else str(value))
        box.press("Tab")

    # Required + Must Pass toggles of the entry row (both are ON by default in Advantage)
    req_note = set_toggles(page, n, row)

    if DRY_RUN:
        shot = f"logs/dryrun_row{row['id']}_grid.png"
        page.screenshot(path=shot)
        return f"grid row [{n}] filled, NOT saved; {req_note}; screenshot {shot}"

    # The row button both recordings clicked before Save (adds the row to the list)
    add_row_to_list(page, n, row["name"])
    return f"{created}{req_note}"


ROW_BTN_INDEX = {"i": None}   # which button in the entry row really adds the row (learned once)


def add_row_to_list(page, n, name):
    """Click the entry row's add button and CHECK the row really moved into the list
    (one more row with this name, entry row empty again). If not, try the next button."""
    before = names_in_grid(page)[name.lower()]
    entry = page.locator(f'input[name="gradeBookComponentTypeId[{GRID_ROW}]"]')
    entry_row = page.locator("tr").filter(has=page.locator(f'input[name="sequence[{n}]"]'))
    cands = entry_row.locator(ROW_BUTTON_SELECTOR)
    if not cands.count():
        cands = entry_row.locator("button")
    order = list(range(cands.count()))
    if ROW_BTN_INDEX["i"] is not None and ROW_BTN_INDEX["i"] in order:
        order.remove(ROW_BTN_INDEX["i"])
        order.insert(0, ROW_BTN_INDEX["i"])
    else:
        order.reverse()                      # the add button is usually the last one in the row
    for idx in order:
        cands.nth(idx).click()
        page.wait_for_timeout(400)
        added = names_in_grid(page)[name.lower()] > before
        cleared = not entry.count() or entry.input_value().strip().lower() != name.lower()
        if added and cleared:
            ROW_BTN_INDEX["i"] = idx
            return
        if not cleared:
            continue                          # nothing happened, try the next button
        # the row was cleared but NOT added (wrong button) -> stop, don't lose data silently
        raise RuntimeError(f"Row '{name}' was cleared but not added to the list - wrong button")
    raise RuntimeError(f"Could not add '{name}' to the grade book list (no working add button)")


def set_toggles(page, n, row):
    """Set Required / Must Pass switches on the entry row to the Excel values (Yes = on)."""
    entry = page.locator("tr").filter(has=page.locator(f'input[name="gradeBookComponentTypeId[{n}]"]'))
    boxes = entry.locator('input[type="checkbox"]')
    notes = []
    if boxes.count() >= 2:
        pairs = (("required", boxes.nth(0), row["required"]), ("mustPass", boxes.nth(1), row["must_pass"]))
    else:
        req = page.locator('input[name="required"]')
        mp = page.locator('input[name="mustPass"], input[name="mustpass"], input[name="isMustPass"]')
        pairs = []
        if req.count():
            pairs.append(("required", req.first, row["required"]))
        if mp.count():
            pairs.append(("mustPass", mp.first, row["must_pass"]))
    for label, box, want in pairs:
        if (box.get_attribute("type") or "").lower() == "checkbox":
            if box.is_checked() != want:
                box.click(force=True)
            notes.append(f"{label}={'Yes' if box.is_checked() else 'No'}")
    return ", ".join(notes) or "toggles: not found"


def option_exists(page, name):
    """Check (without selecting) if a component is in the grid's dropdown list."""
    comp = page.locator(f'input[name="gradeBookComponentTypeId[{GRID_ROW}]"]')
    comp.click()
    comp.fill(name)
    exact = page.get_by_role("option", name=name, exact=True)
    no_options = page.locator(".MuiAutocomplete-noOptions:visible")
    loading = page.locator(".MuiAutocomplete-loading:visible")
    found, end = False, time.time() + 6
    while time.time() < end:            # stop as soon as the list answers (found / "No options")
        if exact.count():
            found = True
            break
        if no_options.count() and not loading.count():
            break
        page.wait_for_timeout(150)
    comp.press("Escape")
    comp.fill("")
    return found


def create_missing_first(page, course_rows, log):
    """Before adding rows to a course: create every component that is not in the list yet,
    so the grid rows can then be added in batches with ONE Save.
    Components already known (state file) are not searched again."""
    known = STATE.get("components", set())
    missing = []
    for r in course_rows:
        if r["name"].lower() in known:
            continue
        if option_exists(page, r["name"]):
            state_add("components", r["name"])
        else:
            missing.append(r)
    if missing:
        say(f"{len(missing)} component(s) not in the list yet - creating them first")
    for r in missing:
        t = f"{datetime.datetime.now():%H:%M:%S}"
        create_component(page, r)
        state_add("components", r["name"])
        log.writerow([t, r["id"], r["course"], r["name"], "CREATED", "New Component saved (not in course yet)"])
        say(f"NEW   #{r['id']} {r['name']} (component created)")


class NeedsSaveFirst(Exception):
    pass


def page_error_text(page):
    """Visible error messages (red alerts / required-field errors) after a Save."""
    msgs = []
    for sel in ('.MuiAlert-standardError:visible', '.MuiAlert-filledError:visible',
                '.MuiSnackbarContent-message:visible', '.Mui-error.MuiFormHelperText-root:visible'):
        loc = page.locator(sel)
        for i in range(min(loc.count(), 3)):
            t = loc.nth(i).inner_text().strip()
            if t:
                msgs.append(t)
    return "; ".join(msgs)


def save_grid(page):
    page.get_by_role("button", name="Save").nth(1).click()
    confirm = page.get_by_role("button", name="Confirm")
    try:
        confirm.first.wait_for(timeout=8000)
        confirm.first.click()
    except PWTimeout:
        err = page_error_text(page)
        raise RuntimeError("Save was not accepted" + (f": {err}" if err else " (no Confirm box appeared)"))
    page.wait_for_timeout(1500)
    err = page_error_text(page)
    if err and "success" not in err.lower():
        raise RuntimeError(f"Advantage showed an error after Save: {err}")


# ------------------------------------------------------------------ data
def load_rows():
    df = pd.read_excel(EXCEL_FILE, sheet_name=SHEET)
    df = df[df["Run"].astype(str).str.strip().str.lower() == "yes"]
    df = df[df["Row ID"] >= START_FROM_ROW_ID]

    def num(v):
        return None if pd.isna(v) else (int(v) if float(v).is_integer() else v)

    def text(r, col):
        return "" if col not in r or pd.isna(r[col]) else str(r[col]).strip()

    def default_course_list(r):   # 'Subject - Program'  ->  'Subject [Program]'
        course, subject = str(r["Course (open in Advantage)"]).strip(), str(r["Subject"]).strip()
        program = course[len(subject) + 3:] if course.startswith(subject + " - ") else course
        return f"{subject} [{program}]"

    rows = []
    for _, r in df.iterrows():
        rows.append(dict(
            id=int(r["Row ID"]), subject=str(r["Subject"]).strip(),
            course=str(r["Course (open in Advantage)"]).strip(),
            name=str(r["Component Name (Description)"]).strip(),
            type=str(r["System Component Type"]).strip(), status=str(r["Status"]).strip(),
            campus=str(r["Campus"]).strip(), number=num(r["Number"]), weight=num(r["Weight"]),
            min=num(r["Min Score"]), seq=num(r["Seq"]),
            required=str(r["Required"]).strip().lower() == "yes",
            must_pass=str(r["Must Pass"]).strip().lower() == "yes",
            course_list=text(r, "Course Name (list)") or default_course_list(r),
            course_min=num(r["Course Min Score"]) if "Course Min Score" in r and not pd.isna(r["Course Min Score"]) else None,
            eff_date=text(r, "Grade Book Effective Date") or "01/01/2020"))
    return rows


# ------------------------------------------------------------------ main
def reopen_grade_book(page, row):
    """Click the grade book again so the grid reloads from the server."""
    gb = page.get_by_role("cell", name=row["course"], exact=True)
    try:
        gb.first.wait_for(timeout=5000)
        gb.first.click()
        grid_inputs(page).first.wait_for(timeout=20000)
        page.wait_for_timeout(800)
    except PWTimeout:
        pass   # grade book list not visible: keep the current grid


def verify_all(page, all_rows, log, old_done):
    """VERIFY_ONLY: open every course, read what exists, compare with the Excel.
    Writes a report and rebuilds the done-cache from what really exists."""
    found, report = [], []
    failed_courses = []
    courses = []
    for r in all_rows:
        if r["course"] not in [c["course"] for c in courses]:
            courses.append(r)
    for crow in courses:
        rows_c = [r for r in all_rows if r["course"] == crow["course"]]
        opened = True
        try:
            open_course_readonly(page, crow)
            full = wait_grid_loaded(page)
            names = names_in_grid(page)
        except Exception as e:
            opened = False
            shot = f"logs/verify_error_{len(failed_courses) + 1}.png"
            try:
                page.screenshot(path=shot)
            except Exception:
                pass
            say(f"WARN  could not open {crow['course_list']}: {str(e).splitlines()[0]} -> {shot}")
            failed_courses.append(crow["course_list"])
            full, names = False, Counter()
        for r in rows_c:
            c = names[r["name"].lower()]
            if not opened:
                status = "UNKNOWN (course could not be opened)"
            else:
                status = "FOUND" if c == 1 else ("DUPLICATE x%d" % c if c > 1 else ("MISSING" if full else "UNKNOWN (grid partly loaded)"))
            report.append((r, status))
            if c >= 1:
                found.append(r)
            elif not full and (r["course"].lower(), r["name"].lower()) in old_done:
                found.append(r)          # could not see it (big grid) but it was done before: keep it
        extra = set(names) - {r["name"].lower() for r in rows_c}
        for nm in sorted(extra):
            report.append(({"id": "", "course": crow["course"], "name": nm}, "IN ADVANTAGE, NOT IN EXCEL"))
        say(f"CHECK {crow['course']}: {sum(1 for r in rows_c if names[r['name'].lower()] >= 1)} of {len(rows_c)} found"
            + ("" if full else " (grid partly loaded)"))
    path = f"logs/verify_{STAMP}.csv"
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["row_id", "course", "component", "result"])
        for r, st_ in report:
            w.writerow([r["id"], r["course"], r["name"], st_])
    if failed_courses:
        say(f"NOT changing the done-cache: could not open {', '.join(failed_courses)}")
    else:
        cache_rewrite(found)
    miss = sum(1 for _, st_ in report if st_ == "MISSING")
    dup = sum(1 for _, st_ in report if st_.startswith("DUPLICATE"))
    say(f"VERIFY done: found {len(found)}, missing {miss}, duplicates {dup}. Report: {path}")
    say("The done-cache now matches Advantage. Run again with VERIFY_ONLY = False to add the missing rows.")


def open_course_readonly(page, row):
    select_course(page, row, None, create=False)


def recover(page):
    """Page stuck: reload, log in again if needed, reopen the program."""
    say("RETRY reloading the page...")
    try:
        page.keyboard.press("Escape")
    except Exception:
        pass
    page.reload(wait_until="domcontentloaded", timeout=60000)
    page.wait_for_timeout(2000)
    if page.get_by_role("button", name="Username & password").count():
        login(page)
    open_program(page)


def main():
    if not USER or not PASS:
        sys.exit("SIS_USER / SIS_PASS are missing in your .env file.")
    if IS_LIVE and not DRY_RUN and not ALLOW_LIVE:
        sys.exit("Stopped: SIS_URL is the LIVE site. Use the DEV site, or set ALLOW_LIVE = True on purpose.")

    all_rows = load_rows()
    done = load_done() if SKIP_DONE_FROM_LOGS else set()
    key = lambda r: (r["course"].lower(), r["name"].lower())
    for r in all_rows:
        if r["id"] in REDO_ROW_IDS and key(r) in done:
            done.discard(key(r))
            say(f"REDO  #{r['id']} {r['name']} (removed from the done-cache for this run)")
    rows = [r for r in all_rows if key(r) not in done][:LIMIT]
    n_done = sum(1 for r in all_rows if key(r) in done)
    say(f"{len(all_rows)} row(s) in Excel | {n_done} already done (cache) | "
        f"{len(rows)} to do | DRY_RUN={DRY_RUN} | batch={BATCH_SIZE} | site={'LIVE' if IS_LIVE else 'DEV'}")
    if not rows and not VERIFY_ONLY:
        say("Nothing to do - everything in the Excel is already saved.")
        return

    with sync_playwright() as p, open(LOG_FILE, "w", newline="", encoding="utf-8") as f:
        log = csv.writer(f)
        log.writerow(["time", "row_id", "course", "component", "status", "message"])
        for r in all_rows:
            if key(r) in done:
                log.writerow(["", r["id"], r["course"], r["name"], "SKIPPED", "already saved (done-cache)"])
        kw = {"headless": HEADLESS, "slow_mo": SLOW_MO}
        if BROWSER_CHANNEL:
            kw["channel"] = BROWSER_CHANNEL
        browser = p.chromium.launch(**kw)
        page = browser.new_context().new_page()
        page.set_default_timeout(20000)

        st = {"ok": 0, "skipped": 0, "failed": 0}
        state = {"course": None, "in_grid": Counter(), "full": False}
        pending = []   # (row, time, message): added to the grid, not saved yet

        def open_and_read(row):
            open_course(page, row, log)
            state["course"] = row["course"]
            state["full"] = wait_grid_loaded(page)
            state["in_grid"] = names_in_grid(page)
            say(f"{sum(state['in_grid'].values())} component(s) seen in this course")
            for nm, c in state["in_grid"].items():
                if c > 1:
                    log.writerow(["", "", row["course"], nm, "DUPLICATE", f"appears {c} times in Advantage - remove extras by hand"])
                    say(f"WARN  duplicate in Advantage: '{nm}' x{c} - remove the extra by hand")
            if PRECREATE_MISSING and CREATE_IF_MISSING and not DRY_RUN and BATCH_SIZE > 1:
                todo = [r for r in rows if r["course"] == row["course"] and key(r) not in done
                        and state["in_grid"][r["name"].lower()] == 0]
                create_missing_first(page, todo, log)
            f.flush()

        def flush():
            """ONE Save for all pending rows -> write done-cache at once -> verify -> log."""
            if not pending:
                return
            say(f"SAVE  {len(pending)} row(s)...")
            save_grid(page)
            # reopen the grade book so the grid shows what the SERVER saved (not what we typed)
            reopen_grade_book(page, pending[0][0])
            state["full"] = wait_grid_loaded(page)
            state["in_grid"] = names_in_grid(page)
            saved, missing = [], []
            for pr, pt, pmsg in pending:
                (saved if state["in_grid"][pr["name"].lower()] > 0 else missing).append((pr, pt, pmsg))
            if missing and not state["full"]:
                # big grid did not load completely: cannot prove it, treat as saved but say so
                saved += [(pr, pt, pmsg + "; UNVERIFIED (grid partly loaded)") for pr, pt, pmsg in missing]
                missing = []
            cache_add([pr for pr, _, _ in saved])
            for pr, pt, pmsg in saved:
                done.add(key(pr))
                st["ok"] += 1
                log.writerow([pt, pr["id"], pr["course"], pr["name"], "OK", f"saved + verified (batch of {len(pending)}); {pmsg}"])
                say(f"OK    #{pr['id']} {pr['name']}")
            for pr, pt, pmsg in missing:
                log.writerow([pt, pr["id"], pr["course"], pr["name"], "NOT SAVED", "not in the grade book after Save"])
                say(f"MISS  #{pr['id']} {pr['name']} - NOT in the grade book after Save")
            pending.clear()
            f.flush()
            if missing:
                raise RuntimeError(f"{len(missing)} row(s) were not saved by Advantage - check the values in the Excel (weight, min score, sequence)")

        try:
            login(page)
            open_program(page)
            if VERIFY_ONLY:
                verify_all(page, all_rows, log, load_done())
                return
            i, attempts = 0, 0
            while i < len(rows):
                row = rows[i]
                t = f"{datetime.datetime.now():%H:%M:%S}"
                try:
                    if key(row) in done:
                        i += 1
                        continue
                    if row["course"] != state["course"]:
                        flush()
                        open_and_read(row)
                    if state["in_grid"][row["name"].lower()] > 0 or any(key(pr) == key(row) for pr, _, _ in pending):
                        st["skipped"] += 1
                        cache_add([row]) if state["in_grid"][row["name"].lower()] > 0 else None
                        done.add(key(row))
                        log.writerow([t, row["id"], row["course"], row["name"], "SKIPPED", "already in course"])
                        say(f"SKIP  #{row['id']} {row['name']} (already there)")
                        i += 1
                        continue
                    try:
                        msg = fill_grid_row(page, row, pending)
                    except NeedsSaveFirst:
                        flush()
                        msg = fill_grid_row(page, row, pending)
                    if DRY_RUN:
                        st["ok"] += 1
                        log.writerow([t, row["id"], row["course"], row["name"], "DRY-RUN", msg])
                        say(f"DRY   #{row['id']} {row['name']}  ({msg})")
                        break
                    pending.append((row, t, msg))
                    say(f"ADD   #{row['id']} {row['name']}  ({len(pending)}/{BATCH_SIZE} in batch)")
                    if len(pending) >= BATCH_SIZE or msg.startswith("created"):
                        flush()
                    i += 1
                    attempts = 0
                    f.flush()
                except DryRunStop as e:
                    st["ok"] += 1
                    log.writerow([t, row["id"], row["course"], row["name"], "DRY-RUN", str(e)])
                    say(f"DRY   {e}")
                    break
                except Exception as e:
                    err = str(e).splitlines()[0][:250]
                    shot = f"logs/error_row{row['id']}.png"
                    try:
                        page.screenshot(path=shot)
                    except Exception:
                        pass
                    for pr, pt, pmsg in pending:
                        if key(pr) in done:  # saved before the problem happened
                            st["ok"] += 1
                            log.writerow([pt, pr["id"], pr["course"], pr["name"], "OK", "saved; check was interrupted"])
                    lost = sum(1 for pr, _, _ in pending if key(pr) not in done)
                    pending.clear()          # unsaved ones will be added again after the retry
                    attempts += 1
                    if attempts <= MAX_RETRIES:
                        say(f"WARN  #{row['id']} {row['name']}: {err}")
                        say(f"RETRY {attempts}/{MAX_RETRIES}" + (f" ({lost} unsaved row(s) will be added again)" if lost else ""))
                        log.writerow([t, row["id"], row["course"], row["name"], "RETRY", err])
                        state["course"] = None
                        # go back to the first row that is not saved yet
                        i = next((k for k, r in enumerate(rows) if key(r) not in done), len(rows))
                        try:
                            recover(page)
                        except Exception as e2:
                            say(f"WARN  reload failed: {str(e2).splitlines()[0]}")
                        continue
                    st["failed"] += 1
                    log.writerow([t, row["id"], row["course"], row["name"], "ERROR", err])
                    say(f"ERROR #{row['id']} {row['name']}: {err}  -> {shot}")
                    say("Stopped after retries. Run again later - saved rows are skipped.")
                    break
            if not st["failed"] and pending:
                try:
                    flush()
                except Exception as e:
                    st["failed"] += 1
                    page.screenshot(path="logs/error_last_batch.png")
                    say(f"ERROR while saving the last batch: {str(e).splitlines()[0]} -> logs/error_last_batch.png")
        finally:
            left = sum(1 for r in all_rows if key(r) not in done)
            say(f"Done. OK={st['ok']} skipped={st['skipped']} errors={st['failed']} | still to do: {left} | Log: {LOG_FILE}")
            if left:
                say("Run again to continue - saved rows are skipped automatically.")
            if DRY_RUN:
                input("Dry run finished. Look at the browser, then press Enter to close it...")
            try:
                browser.close()
            except BaseException:
                pass


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nStopped with Ctrl + C. Run again to continue - finished rows are skipped.")
    except RuntimeError as e:
        print(f"\nSTOPPED: {e}")
