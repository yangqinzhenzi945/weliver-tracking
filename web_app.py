"""Local Streamlit UI for Weliver dispatch status lookup."""
import io
import re
import time
from pathlib import Path

import streamlit as st
from openpyxl import load_workbook
from playwright.sync_api import sync_playwright
from weliver_tracking import URL, open_filter, fill_tracking_number, inspect_result

st.set_page_config(page_title="Weliver 出库查询", page_icon="🚚", layout="centered")
st.title("🚚 Weliver 出库查询")
st.caption("上传 Excel → 在弹出的浏览器登录 Weliver → 自动检查 → 下载结果")
st.info("本工具运行在你自己的电脑上，登录信息和订单不会上传到第三方云端。首次查询请先使用 2–3 个已知订单验证绿色状态点检测。")

file = st.file_uploader("选择出库查询 Excel（.xlsx）", type=["xlsx"])
if file is None:
    st.stop()

try:
    raw = file.getvalue()
    wb = load_workbook(io.BytesIO(raw))
    ws = next((s for s in wb if s.title.strip().lower() == "weliver"), None)
    if ws is None:
        st.error("找不到名为 weliver 的 Sheet，请检查工作表名称。")
        st.stop()
    orders = [
        (row, str(ws.cell(row, 1).value).strip())
        for row in range(1, ws.max_row + 1)
        if ws.cell(row, 1).value
        and re.fullmatch(r"WOSO[\w-]+", str(ws.cell(row, 1).value).strip(), re.I)
    ]
    tasks = [
        (row, no) for row, no in orders
        if str(ws.cell(row, 3).value or "").strip() != "已发出"
    ]
except Exception as exc:
    st.error(f"Excel 无法读取：{exc}")
    st.stop()

a, b, c = st.columns(3)
a.metric("WOSO 总数", len(orders))
b.metric("已发出（跳过）", len(orders) - len(tasks))
c.metric("等待查询", len(tasks))

if not orders:
    st.warning("weliver 工作表 A 列没有找到 WOSO 订单号。")
    st.stop()

st.write("**判定规则：** 运输列表的 Consignment state 至少出现一个绿色圆点，就把该行 C 列写为「已发出」。其他行的 C 列保持不变。")
st.caption("查询过程会弹出 Chromium 浏览器。请手动登录并进入 Transport tracking 列表，程序检测到列表后会自动开始。查询期间不要关闭该浏览器。")

if "result" not in st.session_state:
    st.session_state.result = None
if st.session_state.get("input_name") != file.name or st.session_state.get("input_hash") != hash(raw):
    st.session_state.result = None
    st.session_state.input_name = file.name
    st.session_state.input_hash = hash(raw)

if st.button("开始查询", type="primary", use_container_width=True, disabled=not tasks):
    progress = st.progress(0, text="正在打开 Weliver 登录页面…")
    log = st.empty()
    sent = 0
    errors = []
    unchanged = 0
    try:
        with sync_playwright() as pw:
            context = pw.chromium.launch_persistent_context(
                user_data_dir=str(Path.home() / ".weliver_browser_profile"),
                headless=False,
                viewport={"width": 1440, "height": 900},
            )
            try:
                page = context.pages[0] if context.pages else context.new_page()
                page.goto(URL, wait_until="domcontentloaded", timeout=30000)
                progress.progress(0, text="请在弹出的浏览器登录，并进入 Transport tracking（最多等待 5 分钟）")
                deadline = time.monotonic() + 300
                ready = False
                while time.monotonic() < deadline:
                    if page.is_closed():
                        raise RuntimeError("浏览器已关闭")
                    try:
                        if ("tracking" in page.url.lower() and
                                ("TRANSPORTS FOUND" in page.locator("body").inner_text(timeout=1500).upper() or
                                 "LIST OF TRANSPORTS" in page.locator("body").inner_text(timeout=1500).upper())):
                            ready = True
                            break
                    except Exception:
                        pass
                    page.wait_for_timeout(1200)
                if not ready:
                    raise RuntimeError("没有检测到 Transport tracking 运输列表，请重试。")

                page.wait_for_timeout(1000)
                for i, (row, number) in enumerate(tasks, 1):
                    try:
                        open_filter(page)
                        fill_tracking_number(page, number)
                        found, green, detail = inspect_result(page)
                        if found and green:
                            ws.cell(row, 3).value = "已发出"
                            sent += 1
                            result_label = "已发出"
                        else:
                            unchanged += 1
                            result_label = "保持原值"
                            if not found:
                                errors.append(f"{number}：未能确认查询结果（{detail}）")
                        log.info(f"{i}/{len(tasks)} · {number} · {result_label} · {detail}")
                    except Exception as ex:
                        errors.append(f"{number}：{type(ex).__name__}: {str(ex)[:120]}")
                        log.warning(f"{i}/{len(tasks)} · {number} · 查询失败，未修改")
                    progress.progress(i / len(tasks), text=f"正在查询 {i}/{len(tasks)}")
                    time.sleep(0.2)
            finally:
                context.close()

        result = io.BytesIO()
        wb.save(result)
        st.session_state.result = result.getvalue()
        st.session_state.stats = (sent, unchanged, len(errors))
        st.session_state.issues = errors
        st.success(f"查询结束：新增已发出 {sent} 条，保持原值 {unchanged} 条，需核对 {len(errors)} 条。")
    except Exception as exc:
        st.error(f"查询未完成：{exc}。Excel 未被修改。")

if st.session_state.result is not None:
    stats = st.session_state.stats
    st.success(f"上次结果：新增已发出 {stats[0]} 条，保持原值 {stats[1]} 条，需核对 {stats[2]} 条。")
    name = Path(file.name).stem + "_已更新.xlsx"
    st.download_button(
        "⬇️ 下载查询结果 Excel", data=st.session_state.result,
        file_name=name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary", use_container_width=True
    )
    if st.session_state.issues:
        with st.expander("查看需要人工核对的订单"):
            st.text("\n".join(st.session_state.issues))
