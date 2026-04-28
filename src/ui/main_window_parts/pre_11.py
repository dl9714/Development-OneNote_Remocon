# -*- coding: utf-8 -*-
from __future__ import annotations

from src.ui.main_window_parts._context import (
    bind_context as _bind_context,
    publish_context as _publish_context,
)

_bind_context(globals())



def _resolve_favorite_activation_target(
    target: Dict[str, Any], display_name: str
) -> Dict[str, Any]:
    notebook_text = str((target or {}).get("notebook_text") or "").strip()
    section_text = str((target or {}).get("section_text") or "").strip()
    result = {
        "ok": True,
        "target_kind": None,
        "expected_center_text": "",
        "expected_notebook_text": "",
        "resolved_name": "",
        "resolved_notebook_id": "",
        "error": "",
    }

    if IS_MACOS and section_text:
        result["target_kind"] = "section"
        result["expected_center_text"] = section_text
        result["expected_notebook_text"] = _strip_stale_favorite_prefix(notebook_text)
        return result

    if notebook_text:
        notebook_info = _resolve_notebook_target_for_activation(target, display_name)
        result["target_kind"] = "notebook"
        result["resolved_name"] = notebook_info.get("resolved_name") or notebook_text
        result["resolved_notebook_id"] = notebook_info.get("notebook_id") or ""
        if notebook_info.get("should_abort"):
            result["ok"] = False
            result["error"] = notebook_info.get("error") or ""
            return result
        result["expected_center_text"] = result["resolved_name"] or notebook_text
        return result

    if section_text:
        result["target_kind"] = "section"
        result["expected_center_text"] = section_text
        return result

    result["target_kind"] = "notebook"
    result["expected_center_text"] = _strip_stale_favorite_prefix(display_name)
    return result


def _open_notebook_shortcut_via_shell(
    shortcut_path: str,
    url: str,
    expected_name: str,
    progress_callback: Optional[Callable[[str], None]] = None,
) -> Dict[str, Any]:
    expected_key = _normalize_notebook_name_key(expected_name)
    protocol_url = _build_onenote_protocol_url(shortcut_path, url, expected_name)
    try:
        open_keys = {
            _normalize_notebook_name_key(name)
            for name in _get_open_notebook_names_via_com(refresh=True)
            if _normalize_notebook_name_key(name)
        }
    except Exception:
        open_keys = set()

    if expected_key and expected_key in open_keys:
        return {"ok": True, "already": True, "name": expected_name, "error": ""}

    launch_errors = []
    if progress_callback:
        try:
            progress_callback(f"OneNote 실행 요청 중... {expected_name}")
        except Exception:
            pass

    if protocol_url:
        exe_path = _get_onenote_exe_path()
        _clear_open_notebook_records_cache()
        try:
            if IS_MACOS:
                open_url_in_system(protocol_url)
            elif exe_path:
                subprocess.Popen([exe_path, "/hyperlink", protocol_url])
            else:
                os.startfile(protocol_url)
        except Exception as e:
            launch_errors.append(str(e))
            try:
                if IS_MACOS:
                    open_url_in_system(protocol_url)
                elif exe_path:
                    os.startfile(protocol_url)
                else:
                    _run_powershell(
                        f"Start-Process -FilePath {_ps_quote(protocol_url)}",
                        timeout=10,
                    )
            except Exception as e2:
                launch_errors.append(str(e2))
    else:
        launch_errors.append("OneNote 앱 프로토콜 URL 생성 실패")

    for _ in range(20):
        time.sleep(0.5)
        wait_round = _ + 1
        if progress_callback and (wait_round == 1 or wait_round % 4 == 0):
            try:
                progress_callback(
                    f"OneNote 응답 대기 중... {expected_name} ({wait_round / 2:.1f}초)"
                )
            except Exception:
                pass
        try:
            open_keys = {
                _normalize_notebook_name_key(name)
                for name in _get_open_notebook_names_via_com(refresh=True)
                if _normalize_notebook_name_key(name)
            }
        except Exception:
            continue
        if expected_key and expected_key in open_keys:
            return {
                "ok": True,
                "already": False,
                "name": expected_name,
                "error": "",
            }

    return {
        "ok": False,
        "already": False,
        "name": expected_name,
        "error": "; ".join(msg for msg in launch_errors if msg) or "OneNote 앱에서 열기 실패",
    }


# ----------------- 9. 요소를 중앙으로 위치시키는 함수(최적화) - ensure 호출 -----------------
def _center_element_in_view(
    element_to_center,
    scroll_container,
    *,
    anchor_element=None,
    placement: str = "center",
):
    ensure_pywinauto()
    if not _pwa_ready:
        return
    try:
        if IS_WINDOWS:
            rect_container = _safe_rectangle(scroll_container)
            rect_item = _safe_rectangle(element_to_center)
            rect_anchor = _safe_rectangle(anchor_element) if anchor_element else None
            if _is_already_well_placed_in_view(
                rect_container,
                rect_item,
                rect_anchor,
                placement=placement,
            ):
                return

        try:
            element_to_center.iface_scroll_item.ScrollIntoView()
        except AttributeError:
            return

        settle_timeout = 0.03 if placement == "upper" else 0.1
        settle_interval = 0.01 if placement == "upper" else 0.015
        _wait_rect_settle(
            lambda: element_to_center.rectangle(),
            timeout=settle_timeout,
            interval=settle_interval,
        )

        effective_container = (
            _find_scrollable_ancestor(element_to_center) or scroll_container
        )

        def _anchor_metrics(rect_container, rect_item, rect_anchor):
            if rect_container is None or rect_item is None:
                return rect_item, 0.0, "full"

            if rect_anchor is not None:
                anchor_height = max(1, rect_anchor.bottom - rect_anchor.top)
                if 14 <= anchor_height <= 140:
                    return (
                        rect_anchor,
                        (rect_anchor.top + rect_anchor.bottom) / 2,
                        "anchor_element",
                    )

            container_height = max(1, rect_container.bottom - rect_container.top)
            item_height = max(1, rect_item.bottom - rect_item.top)
            row_height = max(44, min(88, int(container_height * 0.055)))
            if item_height > max(220, int(container_height * 0.85)):
                anchor_bottom = min(rect_item.bottom, rect_item.top + row_height)
                anchor = _make_rect_proxy(
                    rect_item.left,
                    rect_item.top,
                    rect_item.right,
                    anchor_bottom,
                )
                return anchor, (rect_item.top + anchor_bottom) / 2, "top_slice"

            return rect_item, (rect_item.top + rect_item.bottom) / 2, "full"

        def _calc_offset():
            rect_container = _safe_rectangle(effective_container)
            rect_item = _safe_rectangle(element_to_center)
            rect_anchor = _safe_rectangle(anchor_element) if anchor_element else None
            if rect_container is None or rect_item is None:
                return None, None, None, 0.0, "full"
            anchor_rect, item_center_y, anchor_mode = _anchor_metrics(
                rect_container, rect_item, rect_anchor
            )
            container_height = max(1, rect_container.bottom - rect_container.top)
            top_bias = min(48, max(0, int(container_height * 0.08)))
            bottom_bias = min(24, max(0, int(container_height * 0.04)))
            visible_top = rect_container.top + top_bias
            visible_bottom = rect_container.bottom - bottom_bias
            if placement == "upper":
                anchor_top = anchor_rect.top if anchor_rect is not None else rect_item.top
                target_y = visible_top + min(28, max(10, int(container_height * 0.03)))
                offset = anchor_top - target_y
            else:
                container_center_y = (visible_top + visible_bottom) / 2
                offset = item_center_y - container_center_y
            return (
                rect_container,
                rect_item,
                anchor_rect,
                offset,
                anchor_mode,
            )

        rect_container, rect_item, anchor_rect, offset, anchor_mode = _calc_offset()
        debug_hotpaths = _debug_hotpaths_enabled()
        if debug_hotpaths:
            print(
                "[DBG][CENTER][GEOM]",
                f"phase=initial",
                f"placement={placement}",
                f"anchor={anchor_mode}",
                f"offset={offset:.1f}",
                f"container={rect_container}",
                f"item={rect_item}",
                f"anchor_rect={anchor_rect}",
            )

        if _is_already_well_placed_in_view(
            rect_container, rect_item, anchor_rect, placement=placement
        ):
            return

        if abs(offset) <= 10:
            return

        def step_for(dy):
            item_height = 28
            if anchor_rect is not None:
                item_height = min(96, max(20, anchor_rect.bottom - anchor_rect.top))
            elif rect_item is not None:
                item_height = min(96, max(20, rect_item.bottom - rect_item.top))
            return max(1, min(8, int(abs(dy) / max(item_height, 20))))

        max_loops = 3 if placement == "upper" else 6
        for _ in range(max_loops):
            if abs(offset) <= 10:
                break

            direction = "down" if offset > 0 else "up"
            repeats = step_for(offset)

            used_pattern = _scroll_vertical_via_pattern(
                scroll_container, direction=direction, small=True, repeats=repeats
            )
            if not used_pattern:
                wheel_steps = -repeats if offset > 0 else repeats
                _safe_wheel(scroll_container, wheel_steps)

            time.sleep(0.005 if placement == "upper" else 0.01)

            (
                rect_container,
                rect_item,
                anchor_rect,
                offset,
                anchor_mode,
            ) = _calc_offset()

            if _is_already_well_placed_in_view(
                rect_container, rect_item, anchor_rect, placement=placement
            ):
                break

        if debug_hotpaths:
            print(
                "[DBG][CENTER][GEOM]",
                f"phase=final",
                f"placement={placement}",
                f"anchor={anchor_mode}",
                f"offset={offset:.1f}",
                f"container={rect_container}",
                f"item={rect_item}",
                f"anchor_rect={anchor_rect}",
            )

    except Exception as e:
        print(f"[WARN] 중앙 정렬 중 오류: {e}")

_publish_context(globals())
