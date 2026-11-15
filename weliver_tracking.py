"""Weliver tracking -> Excel status updater.

Reads WOSO numbers in column A of the 'weliver' worksheet and writes
'已发出' to column C if at least one green dot appears in the
Consignment state column.

Install: py -m pip install -r requirements.txt
         py -m playwright install chromium
Run:     py weliver_tracking.py "your_file.xlsx"

NOTE: The UI locators are based on screenshots and require live validation.
"""
import argparse
import re
import time
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

URL = "https://weborder.active-logistics.com/customerportal/tracking/list"


def find_filter_button(page):
    buttons = page.locator("button, a, [role='button']")
    for i in range(min(buttons.count(), 120)):
        btn = buttons.nth(i)
        if not btn.is_visible():
            continue
        desc = " ".join(filter(None, [
            btn.get_attribute(x) for x in
            ("title", "aria-label", "data-original-title")
        ]))
        if re.search(r"filter", desc, re.I):
            return btn
    return None


def open_filter(page):
    field = page.get_by_text("Transportnumbers", exact=True)
    if field.count() and field.first.is_visible():
        return
    btn = find_filter_button(page)
    if btn:
        btn.click()
    else:
        page.locator(
            "button:has(i.fa-filter), a:has(i.fa-filter), "
            "[class*='filter']:visible"
        ).first.click(timeout=5000)
    page.get_by_text("Transportnumbers", exact=True).wait_for(
        state="visible", timeout=10000
    )


def fill_tracking_number(page, tracking_number):
    field = page.locator(
        "input[placeholder*='transport' i], "
        "input[name*='transport' i], input[id*='transport' i]"
    )
    if field.count() > 0:
        field.first.fill(tracking_number)
    else:
        label = page.get_by_text("Transportnumbers", exact=True).first
        label.locator("xpath=following::input[1]").fill(tracking_number)
    page.get_by_text("Execute filter", exact=False).last.click()


def inspect_result(page):
    """Return (found, green_dot_detected, diagnostic).

    Green indicator must appear inside the Consignment state table cell.
    Validate this against the real portal before bulk processing.
    """
    try:
        page.wait_for_load_state("networkidle", timeout=9000)
    except PlaywrightTimeoutError:
        page.wait_for_timeout(1400)

    return page.evaluate(r"""() => {
      const normalize = s => (s || '').replace(/\\s+/g, ' ').trim().toLowerCase();
      const text = document.body.innerText || '';
      if (/\\b0\\s+transports?\\s+found\\b/i.test(text))
        return [false, false, '0 transports found'];
      const tables = [...document.querySelectorAll('table')];
      for (const table of tables) {
        const headers = [...table.querySelectorAll('thead th')];
        const index = headers.findIndex(el =>
          normalize(el.textContent).includes('consignment state'));
        if (index < 0) continue;
        const rows = [...table.querySelectorAll('tbody tr')]
          .filter(tr => tr.getClientRects().length);
        for (const tr of rows) {
          const cells = [...tr.children].filter(el => /^(TD|TH)$/.test(el.tagName));
          if (!cells[index]) continue;
          const cell = cells[index];
          const nodes = [cell, ...cell.querySelectorAll('*')];
          const green = nodes.some(el => {
            const r = el.getBoundingClientRect();
            if (r.width < 5 || r.height < 5 || r.width > 25 || r.height > 25)
              return false;
            if (Math.abs(r.width - r.height) > 5) return false;
            const st = getComputedStyle(el);
            const radius = parseFloat(st.borderTopLeftRadius) || 0;
            if (radius < Math.min(r.width, r.height) * 0.3 &&
                el.tagName.toLowerCase() !== 'circle') return false;
            return [st.backgroundColor, st.fill].some(c => {
              const m = /rgba?\\((\\d+)[, ]+([\\d]+)[, ]+([\\d]+)/i.exec(c || '');
              if (!m) return false;
              const [red, gr, blue] = m.slice(1).map(Number);
              return gr > 75 && gr > red * 1.18 &&
                     gr > blue * 1.12 && gr - red > 20;
            });
          });
          return [true, green, green ? 'green dot found' : 'no green dot detected'];
        }
      }
      return [false, false, 'results table not detected'];
    }""")


def main():
    ap = argparse.ArgumentParser(
        description="Update weliver!C from tracking state green dots"
    )
    ap.add_argument("excel", help="Path to Excel workbook")
    ap.add_argument("--headless", action="store_true",
                    help="Not recommended for initial use")
    args = ap.parse_args()
    excel = Path(args.excel).expanduser().resolve()
    if not excel.exists():
        ap.error(f"File not found: {excel}")

    wb = load_workbook(excel)
    sheet = next((s for s in wb if s.title.lower() == "weliver"), None)
    if sheet is None:
        ap.error('Worksheet "weliver" not found')

    tasks = [
        (row, str(sheet.cell(row, 1).value).strip())
        for row in range(1, sheet.max_row + 1)
        if sheet.cell(row, 1).value and re.fullmatch(
            r"WOSO[\\w-]+", str(sheet.cell(row, 1).value).strip(), re.I
        )
    ]
    if not tasks:
        ap.error("No WOSO tracking numbers found in column A")

    already_sent = sum(
        str(sheet.cell(row, 3).value or "").strip() == "已发出"
        for row, _ in tasks
    )
    tasks = [
        (row, num) for row, num in tasks
        if str(sheet.cell(row, 3).value or "").strip() != "已发出"
    ]
    output = excel.with_name(excel.stem + "_已更新.xlsx")
    print(f"Already sent: {already_sent}; to query: {len(tasks)}. Output: {output}")
    if not tasks:
        print("Nothing to query.")
        return

    with sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(
            user_data_dir=str(Path.home() / ".weliver_browser_profile"),
            headless=args.headless,
            viewport={"width": 1440, "height": 900},
            slow_mo=100,
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.goto(URL, wait_until="domcontentloaded")
        print("Log in manually and navigate to Transport tracking.")
        input("When the transport list is visible, press Enter to begin... ")
        results = {"sent": 0, "unchanged": 0, "errors": 0}
        for index, (row, number) in enumerate(tasks, 1):
            try:
                open_filter(page)
                fill_tracking_number(page, number)
                found, green, detail = inspect_result(page)
                if found and green:
                    sheet.cell(row, 3).value = "已发出"
                    results["sent"] += 1
                    status = "已发出"
                else:
                    results["unchanged"] += 1
                    status = "未修改"
                print(f"[{index}/{len(tasks)}] {number}: {status} ({detail})")
            except Exception as ex:
                results["errors"] += 1
                print(
                    f"[{index}/{len(tasks)}] {number}: ERROR "
                    f"{type(ex).__name__}: {str(ex)[:150]}"
                )
                try:
                    page.screenshot(
                        path=str(excel.parent / f"weliver_error_{row}.png")
                    )
                except Exception:
                    pass
            if index % 10 == 0:
                wb.save(output)
                print(f"Progress saved: {output}")
            time.sleep(0.3)
        wb.save(output)
        context.close()
    print(f"Done: {results}. Saved to {output}")


if __name__ == "__main__":
    main()
