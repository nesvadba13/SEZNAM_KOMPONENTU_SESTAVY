# -*- coding: utf-8 -*-

import sys
import re
import os
import shutil
import time
from functools import reduce
from typing import List, Set, Tuple, Optional
import threading
import queue

import creopyson
import tkinter as tk
from tkinter import ttk, filedialog
from PIL import Image, ImageTk

DEBUG = False
DEBUG_TO_STDOUT = False

# ── Konfigurace exportu ──────────────────────────────────────────────────────
EXPORT_PDF = True
EXPORT_DXF = True
EXPORT_STEP = True
INCLUDE_ASSEMBLIES = False
# ────────────────────────────────────────────────────────────────────────────


def _debug(msg: str) -> None:
    if DEBUG:
        stream = sys.stdout if DEBUG_TO_STDOUT else sys.stderr
        print(f"[DEBUG] {msg}", file=stream)


class GuiContext:
    def __init__(self, root: tk.Tk, tree: ttk.Treeview, log: tk.Text) -> None:
        self.root = root
        self.tree = tree
        self.log = log
        self.should_stop = False
        self.items: dict[str, str] = {}
        self.queue: "queue.Queue[tuple[str, tuple]]" = queue.Queue()
        self.busy_label: Optional[ttk.Label] = None
        self.run_btn: Optional[ttk.Button] = None
        self.active_asm_label: Optional[ttk.Label] = None
        self.output_dir_label: Optional[ttk.Label] = None
        self.busy_running = False
        self.busy_dots = 0
        self.current_active_model = ""
        self.poll_countdown = 0
        self.export_running = False
        self.output_directory = ""


def _load_logo() -> Optional[ImageTk.PhotoImage]:
    """Nacte a upravi logo pro zobrazeni."""
    try:
        # cesta k logu vzhledem ke skriptu
        script_dir = os.path.dirname(os.path.abspath(__file__))
        logo_path = os.path.join(script_dir, "obrázky", "LC_Lifts.jpg")
        
        if os.path.exists(logo_path):
            # nacti a zmensi logo
            image = Image.open(logo_path)
            # zachovej proporce, max vyska 60px
            image.thumbnail((120, 60), Image.Resampling.LANCZOS)
            return ImageTk.PhotoImage(image)
    except Exception as e:
        _debug(f"Nepodařilo se načíst logo: {e}")
    return None


def _normalize_filename(name: str) -> str:
    parts = name.split(".")
    if len(parts) >= 3 and parts[-1].isdigit():
        return ".".join(parts[:-1])
    return name


def _is_model(name: str, item_type: str | None) -> bool:
    n = name.lower()
    if n.endswith(".prt") or n.endswith(".asm") or ".prt." in n or ".asm." in n:
        return True
    if item_type:
        t = item_type.strip().lower()
        return t in {"prt", "part", "asm", "assembly"}
    return False


def _extract_filename_from_pathlike(s: str) -> str | None:
    s = s.strip()
    tokens = re.split(r'[:/\\\s]+', s)
    for t in reversed(tokens):
        t_norm = _normalize_filename(t)
        if t_norm.lower().endswith((".prt", ".asm")):
            return t_norm
    m = re.search(r'([A-Za-z0-9_\-]+\.((?:prt)|(?:asm))(?:\.\d+)?)', s, flags=re.IGNORECASE)
    if m:
        return _normalize_filename(m.group(1))
    return None


def _iter_bom_nodes(node):
    if isinstance(node, dict):
        yield node
        children = node.get("children")
        if isinstance(children, dict) and "children" in children:
            children = children.get("children")
        if isinstance(children, list):
            for ch in children:
                yield from _iter_bom_nodes(ch)
        elif isinstance(children, dict):
            yield from _iter_bom_nodes(children)
    elif isinstance(node, (list, tuple)):
        for it in node:
            yield from _iter_bom_nodes(it)


def get_all_model_names_recursive(
    client: creopyson.Client,
    asm_file: str | None = None,
    include_assemblies: bool = True,
) -> List[str]:
    if asm_file is None:
        active = client.file_get_active()
        _debug(f"file_get_active() => {active!r}")
        if not active or "file" not in active or not active["file"]:
            raise RuntimeError("V Creo neni aktivni zadny model.")
        asm_file = active["file"]
        work_dir = active.get("dirname") if isinstance(active, dict) else None
        if work_dir:
            _debug(f"Pracovni slozka: {work_dir}")

    _debug(f"Aktivni/pozadovany model (asm_file): {asm_file}")

    # Kontrola zda je aktivni soubor dil (.prt) nebo sestava (.asm)
    asm_file_norm = _normalize_filename(asm_file)
    if asm_file_norm.lower().endswith(".prt"):
        _debug(f"Aktivni je dil {asm_file_norm}, vracim jen tento dil")
        return [asm_file_norm]

    # Pro sestavy pouzij BOM
    try:
        try:
            bom_items = client.bom_get_paths(file_=asm_file, exclude_inactive=True)
        except TypeError:
            _debug("bom_get_paths nepodporuje exclude_inactive, volam bez nej")
            bom_items = client.bom_get_paths(file_=asm_file)
    except Exception as exc:
        raise RuntimeError(f"Selhalo nacteni BOM pro '{asm_file}': {exc}") from exc

    try:
        total = len(bom_items) if hasattr(bom_items, "__len__") else None
    except Exception:
        total = None
    _debug(f"bom_get_paths() typ={type(bom_items).__name__}, pocet={total}")

    models: Set[str] = set()
    node_count = 0
    for node in _iter_bom_nodes(bom_items):
        node_count += 1
        name = None
        i_type = None
        if isinstance(node, dict):
            name = node.get("file") or node.get("filename") or node.get("name")
            if not name:
                path_like = node.get("path") or node.get("component_path") or node.get("pathname")
                if isinstance(path_like, str):
                    name = _extract_filename_from_pathlike(path_like)
            i_type = node.get("type")
        if not name:
            continue

        name = _normalize_filename(name)
        lower_name = name.lower()

        if include_assemblies:
            if _is_model(name, i_type):
                models.add(name)
        else:
            is_part = lower_name.endswith(".prt") or ".prt." in lower_name
            if i_type:
                is_part = is_part or i_type.strip().lower() in {"prt", "part"}
            if is_part:
                models.add(name)

    if include_assemblies and asm_file and asm_file.lower().endswith(".asm"):
        models.add(_normalize_filename(asm_file))

    _debug(f"Prosel jsem {node_count} uzlu BOM stromu; nalezeno_modelu={len(models)}")
    return sorted(models, key=str.lower)


def _get_work_dir_from_active(client: creopyson.Client) -> Optional[str]:
    try:
        active = client.file_get_active()
        if isinstance(active, dict):
            return active.get("dirname")
    except Exception as exc:
        _debug(f"file_get_active() (workdir) vyjimka: {exc!r}")
    return None


def _get_active_asm_from_creo(client: creopyson.Client) -> Optional[str]:
    try:
        active = client.file_get_active()
        if isinstance(active, dict):
            return active.get("file")
    except Exception as exc:
        _debug(f"file_get_active() (active_asm) vyjimka: {exc!r}")
    return None


def _strip_version_suffix(name: str) -> str:
    return _normalize_filename(name)


def _drw_name_for_model(model_name: str) -> str:
    stem = _strip_version_suffix(model_name).rsplit(".", 1)[0]
    return f"{stem}.drw"


def _sanitize_filename(name: str) -> str:
    name = name.replace(" ", "_")
    return re.sub(r'[<>:"/\\|?*]+', "_", name).strip().rstrip(".")


def _open_model(client: creopyson.Client, model_name: str, work_dir: Optional[str]) -> bool:
    try:
        client.file_open(file_=model_name, dirname=work_dir) if work_dir else client.file_open(file_=model_name)
        _debug(f"Otevren model: {model_name}")
        return True
    except Exception as exc:
        _debug(f"file_open(model={model_name}, dir={work_dir}) selhalo: {exc!r}")
        return False


def _open_drawing(client: creopyson.Client, drw_name: str, work_dir: Optional[str]) -> bool:
    try:
        client.file_open(file_=drw_name, dirname=work_dir) if work_dir else client.file_open(file_=drw_name)
        _debug(f"Otevren vykres: {drw_name}")
        return True
    except Exception as exc:
        _debug(f"file_open(drw={drw_name}, dir={work_dir}) selhalo: {exc!r}")
        return False


def _regenerate_model_safe(
    client: creopyson.Client,
    model_name: str,
    gui: Optional[GuiContext] = None,
) -> None:
    try:
        client.file_regenerate(model_name)
        _debug(f"Regenerace modelu: {model_name}")
        _gui_log(gui, f"Regenerace OK: {model_name}")
    except Exception as exc:
        _debug(f"file_regenerate({model_name}) vyjimka: {exc!r}")
        _gui_log(gui, f"Regenerace selhala ({model_name}): {exc}")


def _drawing_regenerate_safe(client: creopyson.Client, drawing_name: str) -> None:
    try:
        client.drawing_regenerate(drawing=drawing_name)
        _debug(f"Regenerace vykresu: {drawing_name}")
    except Exception as exc:
        _debug(f"drawing_regenerate({drawing_name}) vyjimka: {exc!r}")


def _get_cur_model_in_drawing(client: creopyson.Client) -> Optional[str]:
    try:
        return client.drawing_get_cur_model()
    except Exception as exc:
        _debug(f"drawing_get_cur_model() vyjimka: {exc!r}")
        return None


def _get_param_value(client: creopyson.Client, model_name: str, name: str) -> Optional[str]:
    try:
        raw = client.parameter_list(file_=model_name, name=name)
        if isinstance(raw, list) and raw:
            merged = reduce(lambda a, b: dict(a, **b), raw)
            val = merged.get("value")
            return str(val) if val is not None else None
    except Exception as exc:
        _debug(f"parameter_list({model_name}, {name}) vyjimka: {exc!r}")
    return None


def _ensure_output_dir(client: creopyson.Client, custom_dir: str = "") -> str:
    if custom_dir and os.path.exists(custom_dir):
        _debug(f"Pouzivam uzivatelem zvolenou slozku: {custom_dir}")
        return custom_dir
    
    # implicitni chovani - vytvori "DXF+PDF" v pracovnim adresari
    try:
        base = creopyson.creo.pwd(client)
    except Exception as exc:
        _debug(f"creo.pwd vyjimka: {exc!r}")
        base = "."
    target = os.path.join(base, "DXF+PDF")
    if not os.path.exists(target):
        os.makedirs(target, exist_ok=True)
        _debug(f"Vytvorena slozka: {target}")
        time.sleep(0.2)
    else:
        _debug(f"Slozka jiz existuje: {target}")
    return target


def _get_default_output_dir(client: creopyson.Client) -> str:
    """Získá výchozí výstupní složku (DXF+PDF v pracovním adresáři)."""
    try:
        base = creopyson.creo.pwd(client)
    except Exception as exc:
        _debug(f"creo.pwd vyjimka: {exc!r}")
        base = "."
    return os.path.join(base, "DXF+PDF")


def _gui_log(gui: Optional[GuiContext], message: str) -> None:
    if gui is None:
        return
    gui.queue.put(("log", (message,)))


def _gui_update_row(
    gui: Optional[GuiContext],
    model: str,
    pdf_path: str,
    dxf_path: str,
    step_path: str,
) -> None:
    if gui is None:
        return
    gui.queue.put(("row", (model, pdf_path, dxf_path, step_path)))


def _gui_set_busy(gui: Optional[GuiContext], running: bool) -> None:
    if gui is None:
        return
    gui.queue.put(("busy", (running,)))


def _gui_set_active_asm(gui: Optional[GuiContext], name: str) -> None:
    if gui is None:
        return
    gui.queue.put(("active_asm", (name,)))


def _gui_set_output_dir(gui: Optional[GuiContext], path: str) -> None:
    if gui is None:
        return
    gui.queue.put(("output_dir", (path,)))


def _start_busy_indicator(gui: GuiContext) -> None:
    gui.busy_running = True
    gui.busy_dots = 0
    _animate_busy(gui)


def _stop_busy_indicator(gui: GuiContext) -> None:
    gui.busy_running = False
    if gui.busy_label:
        gui.busy_label.config(text="Hotovo.")
    if gui.run_btn:
        gui.run_btn.config(state="normal")


def _animate_busy(gui: GuiContext) -> None:
    if not gui.busy_running:
        return
    gui.busy_dots = (gui.busy_dots + 1) % 4
    dots = "." * gui.busy_dots
    if gui.busy_label:
        gui.busy_label.config(text=f"Pracuji{dots}")
    gui.root.after(500, _animate_busy, gui)


def _update_status_label(gui: GuiContext) -> None:
    """Aktualizuje zobrazeni aktivniho modelu + countdown."""
    if gui.active_asm_label:
        model_part = f"Aktivni model: {gui.current_active_model or '(neznamy)'}"
        if gui.export_running:
            gui.active_asm_label.config(text=f"{model_part} | Export probiha...")
        else:
            countdown_part = f"Obnova za: {gui.poll_countdown} s"
            gui.active_asm_label.config(text=f"{model_part} | {countdown_part}")


def _update_output_dir_label(gui: GuiContext) -> None:
    """Aktualizuje zobrazeni vystupni slozky."""
    if gui.output_dir_label:
        if gui.output_directory:
            # zkraceni cesty pro zobrazeni
            display_path = gui.output_directory
            if len(display_path) > 60:
                display_path = "..." + display_path[-57:]
            gui.output_dir_label.config(text=f"Vystupni slozka: {display_path}")
        else:
            gui.output_dir_label.config(text="Vystupni slozka: (implicitne DXF+PDF)")


def _process_gui_queue(gui: GuiContext) -> None:
    try:
        while True:
            kind, payload = gui.queue.get_nowait()
            if kind == "log":
                (message,) = payload
                gui.log.insert("end", message + "\n")
                gui.log.see("end")
            elif kind == "row":
                model, pdf_path, dxf_path, step_path = payload
                pdf_status = "OK" if pdf_path else "NE"
                dxf_status = "OK" if dxf_path else "NE"
                step_status = "OK" if step_path else "NE"
                item_id = gui.items.get(model)
                if item_id is None:
                    item_id = gui.tree.insert("", "end", values=(model, pdf_status, dxf_status, step_status))
                    gui.items[model] = item_id
                else:
                    gui.tree.item(item_id, values=(model, pdf_status, dxf_path, step_path))
            elif kind == "busy":
                (running,) = payload
                if running:
                    gui.export_running = True
                    _start_busy_indicator(gui)
                else:
                    gui.export_running = False
                    _stop_busy_indicator(gui)
                _update_status_label(gui)
            elif kind == "active_asm":
                (name,) = payload
                gui.current_active_model = name
                _update_status_label(gui)
            elif kind == "output_dir":
                (path,) = payload
                gui.output_directory = path
                _update_output_dir_label(gui)
    except queue.Empty:
        pass
    gui.root.after(100, _process_gui_queue, gui)


def process_drawings_for_models(
    client: creopyson.Client,
    models: List[str],
    work_dir: Optional[str],
    output_dir: str,
    export_pdf: bool,
    export_dxf: bool,
    export_step: bool,
    gui: Optional[GuiContext] = None,
) -> List[Tuple[str, str, str, str]]:
    results: List[Tuple[str, str, str, str]] = []
    out_dir = _ensure_output_dir(client, output_dir)

    _debug(f"Start zpracovani {len(models)} modelu")
    _gui_log(gui, f"Vystupni slozka: {out_dir}")
    ok_pdf = 0
    ok_dxf = 0
    ok_step = 0

    for model in models:
        if gui and gui.should_stop:
            _gui_log(gui, "Zastaveno uzivatelem.")
            break

        model_clean = _strip_version_suffix(model)
        model_stem = model_clean.rsplit(".", 1)[0]
        drawing_name = _drw_name_for_model(model)
        status_msg = f"Zpracovavam: {model_clean} (DRW={drawing_name})"
        print(status_msg)
        _gui_log(gui, status_msg)

        if not _open_model(client, model_clean, work_dir):
            msg = f"{model_clean}: PDF:NE / DXF:NE / STEP:NE (model nelze otevrit)"
            print(msg)
            _gui_log(gui, msg)
            results.append((model, "", "", ""))
            _gui_update_row(gui, model, "", "", "")
            continue

        _gui_log(gui, f"Regeneruji model: {model_clean}")
        _regenerate_model_safe(client, model_clean, gui)

        if not _open_drawing(client, drawing_name, work_dir):
            msg = f"{model_clean}: PDF:NE / DXF:NE / STEP:NE (chybi/neotevren DRW)"
            print(msg)
            _gui_log(gui, msg)
            results.append((model, "", "", ""))
            _gui_update_row(gui, model, "", "", "")
            try:
                client.file_close(file_=model_clean)
            except Exception:
                pass
            continue

        _gui_log(gui, f"Regeneruji vykres: {drawing_name}")
        _drawing_regenerate_safe(client, drawing_name)

        cur_model = _get_cur_model_in_drawing(client) or model_clean
        param_c_vykresu = _get_param_value(client, cur_model, "c_vykresu")
        param_nazev = _get_param_value(client, cur_model, "nazev")

        pdf_path_final = ""
        dxf_path_final = ""
        step_path_final = ""

        sheet_count = 1
        try:
            sheets = creopyson.drawing.list_sheets(client, drawing=drawing_name)
            sheet_count = len(sheets) if isinstance(sheets, list) else 1
            _debug(f"Pocet listu pro {drawing_name}: {sheet_count}")
            _gui_log(gui, f"Pocet listu: {sheet_count}")
        except Exception as exc_sc:
            _debug(f"list_sheets vyjimka, predpokladam 1 list: {exc_sc!r}")

        if export_pdf:
            try:
                if param_c_vykresu and param_nazev:
                    base = _sanitize_filename(f"{param_c_vykresu}-{param_nazev}")
                elif param_c_vykresu:
                    base = _sanitize_filename(f"{param_c_vykresu}-{model_stem}")
                else:
                    base = None

                if base:
                    name_only = base[:30]
                    pdf_filename = f"{name_only}.pdf"
                    creopyson.interface.export_pdf(client, file_=drawing_name, filename=pdf_filename, sheet_range="1")
                    source_path = creopyson.creo.pwd(client) + pdf_filename
                    destination_path = os.path.join(out_dir, pdf_filename)
                    time.sleep(0.5)
                    shutil.move(source_path, destination_path)
                    pdf_path_final = destination_path
            except Exception as e:
                _gui_log(gui, f"PDF Preskoceno ({model_clean}): {e}")
                pdf_path_final = ""

        if export_dxf:
            try:
                if param_c_vykresu:
                    parametar_c_vykresu = _sanitize_filename(param_c_vykresu)
                    dxf_filename = f"{parametar_c_vykresu}.dxf"
                    try:
                        creopyson.drawing.select_sheet(client, "2", drawing=drawing_name)
                    except Exception:
                        pass
                    creopyson.interface.export_file(client, "DXF", file_=drawing_name, filename=parametar_c_vykresu)
                    source_path = creopyson.creo.pwd(client) + dxf_filename
                    destination_path = os.path.join(out_dir, dxf_filename)
                    time.sleep(0.5)
                    shutil.move(source_path, destination_path)
                    dxf_path_final = destination_path
            except Exception as e:
                _gui_log(gui, f"DXF Preskoceno ({model_clean}): {e}")
                dxf_path_final = ""

        if export_step:
            try:
                if param_c_vykresu:
                    safe_step_name_base = _sanitize_filename(param_c_vykresu)
                    step_filename = f"{safe_step_name_base}.stp"
                    creopyson.interface.export_file(client, "STEP", file_=model_clean, filename=safe_step_name_base)
                    source_step = creopyson.creo.pwd(client) + step_filename
                    destination_step = os.path.join(out_dir, step_filename)
                    time.sleep(0.5)
                    shutil.move(source_step, destination_step)
                    step_path_final = destination_step
            except Exception as exc:
                _gui_log(gui, f"STEP preskoceno ({model_clean}): {exc}")
                step_path_final = ""

        try:
            client.file_close(file_=drawing_name)
        except Exception:
            pass
        try:
            client.file_close(file_=model_clean)
        except Exception:
            pass

        results.append((model, pdf_path_final, dxf_path_final, step_path_final))

        pdf_ok = "OK" if pdf_path_final else "NE"
        dxf_ok = "OK" if dxf_path_final else "NE"
        step_ok = "OK" if step_path_final else "NE"
        line = f"{model_clean} - PDF:{pdf_ok} / DXF:{dxf_ok} / STEP:{step_ok}"
        print(line)
        _gui_log(gui, line)
        _gui_update_row(gui, model, pdf_path_final, dxf_path_final, step_path_final)

        if pdf_path_final:
            ok_pdf += 1
        if dxf_path_final:
            ok_dxf += 1
        if step_path_final:
            ok_step += 1

    summary = f"Souhrn: PDF OK={ok_pdf}, DXF OK={ok_dxf}, STEP OK={ok_step}, celkem modelu={len(results)}"
    print(summary)
    _gui_log(gui, summary)

    return results


def _restore_active_assembly(client: creopyson.Client, asm_name: Optional[str]) -> None:
    if not asm_name:
        return
    try:
        client.file_open(file_=asm_name)
    except Exception as exc:
        _debug(f"file_open({asm_name}) vyjimka: {exc!r}")
    try:
        client.file_activate(file_=asm_name)
        _debug(f"Aktivovana vychozi sestava: {asm_name}")
    except Exception as exc:
        _debug(f"file_activate({asm_name}) vyjimka: {exc!r}")


def _poll_active_asm(client: creopyson.Client, gui: GuiContext) -> None:
    """Kazdych 15 sekund aktualizuje zobrazenou aktivni sestavu."""
    if not gui.export_running:
        try:
            name = _get_active_asm_from_creo(client)
            _gui_set_active_asm(gui, name or "")
        except Exception:
            pass
        
        gui.poll_countdown = 5
        _update_status_label(gui)
    
    if not gui.should_stop:
        gui.root.after(5_000, _poll_active_asm, client, gui)


def _countdown_tick(gui: GuiContext) -> None:
    """Kazdou sekundu snizi countdown o 1."""
    if not gui.export_running and gui.poll_countdown > 0:
        gui.poll_countdown -= 1
        _update_status_label(gui)
    
    if not gui.should_stop:
        gui.root.after(1000, _countdown_tick, gui)


def _create_main_window(client: creopyson.Client) -> GuiContext:
    root = tk.Tk()
    root.title("Export vykresu a modelu")
    root.geometry("820x660")

    main_frame = ttk.Frame(root, padding=8)
    main_frame.pack(fill="both", expand=True)

    # ── Horni panel: konfigurace + logo ──────────────────────────────────────
    top_frame = ttk.LabelFrame(main_frame, text="Nastaveni", padding=6)
    top_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 6))
    top_frame.columnconfigure(4, weight=1)

    pdf_var = tk.BooleanVar(value=EXPORT_PDF)
    dxf_var = tk.BooleanVar(value=EXPORT_DXF)
    step_var = tk.BooleanVar(value=EXPORT_STEP)
    asm_var = tk.BooleanVar(value=INCLUDE_ASSEMBLIES)

    ttk.Checkbutton(top_frame, text="PDF", variable=pdf_var).grid(row=0, column=0, sticky="w", padx=(0, 8))
    ttk.Checkbutton(top_frame, text="DXF", variable=dxf_var).grid(row=1, column=0, sticky="w", padx=(0, 8))
    ttk.Checkbutton(top_frame, text="STEP", variable=step_var).grid(row=2, column=0, sticky="w", padx=(0, 8))
    ttk.Checkbutton(top_frame, text="Sestavy (ASM)", variable=asm_var).grid(row=3, column=0, sticky="w", padx=(0, 16))

    # tlacitko pro nastaveni vychozi slozky
    def set_default_output_dir():
        default_dir = _get_default_output_dir(client)
        _gui_set_output_dir(gui, "")  # prazdny string znamena pouzij vychozi
    
    # tlacitko pro vyber vlastni vystupni slozky
    def choose_output_dir():
        folder = filedialog.askdirectory(title="Vyberte vystupni slozku")
        if folder:
            _gui_set_output_dir(gui, folder)
    
    ttk.Button(top_frame, text="Vychozi slozka", command=set_default_output_dir).grid(row=6, column=5, sticky="e", padx=(8, 4))
    ttk.Button(top_frame, text="Vybrat slozku", command=choose_output_dir).grid(row=6, column=4, sticky="e", padx=(4, 0))

    # logo vpravo nahore
    logo = _load_logo()
    if logo:
        logo_label = ttk.Label(top_frame, image=logo)
        logo_label.image = logo  # uchovat referenci
        logo_label.grid(row=0, column=5, rowspan=3, sticky="ne", padx=(8, 0))

    active_asm_label = ttk.Label(top_frame, text="Aktivni model: (neznamy) | Obnova za: 0 s", anchor="w", foreground="gray")
    active_asm_label.grid(row=4, column=0, columnspan=6, sticky="ew", pady=(4, 0))

    output_dir_label = ttk.Label(top_frame, text="Vystupni slozka: (implicitne DXF+PDF)", anchor="w", foreground="blue")
    output_dir_label.grid(row=6, column=0, columnspan=4, sticky="ew")

    # ── Tabulka vysledku ─────────────────────────────────────────────────────
    tree_cols = ("model", "pdf", "dxf", "step")
    tree = ttk.Treeview(main_frame, columns=tree_cols, show="headings", height=10)
    tree.heading("model", text="Model")
    tree.heading("pdf", text="PDF")
    tree.heading("dxf", text="DXF")
    tree.heading("step", text="STEP")
    tree.column("model", width=340, anchor="w")
    tree.column("pdf", width=80, anchor="center")
    tree.column("dxf", width=80, anchor="center")
    tree.column("step", width=80, anchor="center")

    tree_scroll = ttk.Scrollbar(main_frame, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=tree_scroll.set)

    tree.grid(row=1, column=0, sticky="nsew", padx=(0, 4), pady=(0, 4))
    tree_scroll.grid(row=1, column=1, sticky="ns", pady=(0, 4))

    # ── Log ──────────────────────────────────────────────────────────────────
    log = tk.Text(main_frame, height=8, wrap="word")
    log_scroll = ttk.Scrollbar(main_frame, orient="vertical", command=log.yview)
    log.configure(yscrollcommand=log_scroll.set)
    log.grid(row=2, column=0, sticky="nsew", padx=(0, 4))
    log_scroll.grid(row=2, column=1, sticky="ns")

    # ── Spodni lista: stav + tlacitko ────────────────────────────────────────
    bottom_frame = ttk.Frame(main_frame)
    bottom_frame.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(6, 0))
    bottom_frame.columnconfigure(0, weight=1)

    busy_label = ttk.Label(bottom_frame, text="")
    busy_label.grid(row=0, column=0, sticky="w")

    run_btn = ttk.Button(bottom_frame, text="Spustit export")
    run_btn.grid(row=0, column=1, sticky="e", padx=(6, 0))

    main_frame.rowconfigure(1, weight=3)
    main_frame.rowconfigure(2, weight=1)
    main_frame.columnconfigure(0, weight=1)

    gui = GuiContext(root, tree, log)
    gui.busy_label = busy_label
    gui.run_btn = run_btn
    gui.active_asm_label = active_asm_label
    gui.output_dir_label = output_dir_label
    gui.poll_countdown = 15

    # inicializace labelu
    _update_output_dir_label(gui)

    # ── Tlacitko Spustit ─────────────────────────────────────────────────────
    def on_run() -> None:
        for item in tree.get_children():
            tree.delete(item)
        gui.items.clear()
        gui.should_stop = False

        export_pdf = pdf_var.get()
        export_dxf = dxf_var.get()
        export_step = step_var.get()
        include_assemblies = asm_var.get()

        run_btn.config(state="disabled")
        _gui_set_busy(gui, True)

        def worker() -> None:
            try:
                try:
                    running = client.is_creo_running()
                    if not running:
                        msg = "Creo nebezi nebo neni dostupne."
                        print(msg)
                        _gui_log(gui, msg)
                        return
                except Exception as exc2:
                    _debug(f"is_creo_running() vyjimka: {exc2!r}")

                original_active: Optional[str] = None
                try:
                    active_info = client.file_get_active()
                    if isinstance(active_info, dict):
                        original_active = active_info.get("file")
                        _gui_set_active_asm(gui, original_active or "")
                except Exception:
                    pass

                _gui_log(gui, "Nacitam seznam modelu...")
                models = get_all_model_names_recursive(
                    client, asm_file=None, include_assemblies=include_assemblies
                )
                _gui_log(gui, f"Nalezeno modelu: {len(models)}")

                if not models:
                    _gui_log(gui, "Nebyl nalezen zadny model.")
                    return

                work_dir = _get_work_dir_from_active(client)
                if work_dir:
                    try:
                        try:
                            client.creo_cd(work_dir)
                        except Exception:
                            client.file_set_curdir(work_dir)
                    except Exception as exc2:
                        _gui_log(gui, f"Nepodarilo se nastavit pracovni adresar: {exc2}")
                else:
                    _gui_log(gui, "Pracovni slozka neznama, export do aktualniho adresare.")

                _gui_log(gui, "Spoustim export...")
                process_drawings_for_models(
                    client, models, work_dir, gui.output_directory, export_pdf, export_dxf, export_step, gui
                )

                _restore_active_assembly(client, original_active)
                _gui_log(gui, "Hotovo.")
            finally:
                _gui_set_busy(gui, False)
                print("[END]")

        threading.Thread(target=worker, daemon=True).start()

    run_btn.config(command=on_run)

    def on_close() -> None:
        gui.should_stop = True
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.after(100, _process_gui_queue, gui)
    root.after(500, _poll_active_asm, client, gui)
    root.after(1000, _countdown_tick, gui)

    return gui


def main() -> None:
    print("[START]")

    client = creopyson.Client()
    version = getattr(creopyson, "__version__", None)
    if version:
        _debug(f"creopyson verze: {version}")

    try:
        client.connect()
        _debug("Pripojeni k Creoson serveru: OK")
    except Exception as exc:
        print(f"Nelze se pripojit k Creoson serveru (localhost:9056): {exc}")

    gui = _create_main_window(client)
    gui.root.mainloop()
    print("[END]")


if __name__ == "__main__":
    main()
