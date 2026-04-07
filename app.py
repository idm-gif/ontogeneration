#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Sheets → XML Graph Converter
Fixed:
1. Added "Data Separator" to split cell values into multiple data entries (Array support).
2. Applies globally to all data attributes in both Simple and Advanced modes.
3. Implemented "structure" sheet lookup to auto-populate data for parent nodes generated via hierarchy/grouping.
4. FIX: Structure sheet enrichment now correctly applied to ALL parent/phantom/sheet nodes.
5. FIX: Plain image/youtube URLs (no HYPERLINK formula) are now properly typed in auto mode.
6. FIX: Sheet-level nodes and grouping parent nodes now get data from structure sheet on creation.
7. FIX: Grouping now respects per-sheet hierarchy configs — rows are routed through _build_hier
         on the filtered subset, producing Type→Sheet→Phantom(s)→Row topology.
8. NEW: XLSX Ctrl+K hyperlinks and =HYPERLINK() formulas are extracted via openpyxl and
         injected so link/text/type are all preserved correctly.
9. NEW: Google Sheets loaded with value_render_option='FORMULA' so =HYPERLINK() formulas
         are preserved and typed correctly.
10. NEW: "Exclude Title Column from Node Data" checkbox — when on, the column chosen as the
          node title is not added as a separate data element (avoids redundancy). On by default.
"""
import os
import re
import uuid
import json
import traceback
import datetime
import typing as t
from difflib import SequenceMatcher
import xml.etree.ElementTree as ET
from xml.dom import minidom
import pandas as pd
import gradio as gr
 
# --- [BACKEND LOGIC] ---
HAS_SHEETS = False
try:
    import gspread
    from google.oauth2.service_account import Credentials
    HAS_SHEETS = True
except ImportError:
    pass
 
TYPE_CHOICES = ["auto", "text", "link", "image", "youtube"]
DELIMITER_CHOICES = [";", ",", "|", "No Separation"]
 
 
class GraphNode:
    def __init__(self, name: str, data_elements: t.List[dict] = None, node_type: str = "", is_intermediate: bool = False):
        self.guid = uuid.uuid4().hex[:12].upper()
        self.name = str(name).strip() if name else "Untitled"
        self.data_elements = data_elements or []
        self.node_type = node_type
        self.shape = "circle"
        self.x_pos = 0
        self.y_pos = 0
        self.color = "255" if is_intermediate else "13421772"
        self.is_intermediate = is_intermediate
 
 
class GraphEdge:
    def __init__(self, node1: str, node2: str, bilateral: bool = False):
        self.guid = uuid.uuid4().hex[:12].upper()
        self.node1 = str(node1).strip()
        self.node2 = str(node2).strip()
        self.bilateral = bilateral
        self.group = ""
 
 
class Converter:
    def __init__(self):
        self.sheets_data: dict[str, pd.DataFrame] = {}
        self.root_name: str = "Graph"
        self.loaded_sheets: list[str] = []
        self.selected_sheets: set[str] = set()
        self.edge_direction: str = "right"
        self.show_empty_nodes: bool = True
        self.column_type_preferences: dict[str, dict[str, str]] = {}
        self.hierarchy_configs: dict[str, list[dict]] = {}
        self.sheet_title_cols: dict[str, str] = {}
        self.nodes: list[GraphNode] = []
        self.edges: list[GraphEdge] = []
        self.node_map: dict[str, GraphNode] = {}
        self.phantom_nodes: dict = {}
        self.grouping_rules: list[dict] = []
        self.manual_grouping_rules: list[dict] = []
        self.grouping_enabled: bool = False
        self.sheets_in_grouping: set[str] = set()
 
        # Flags
        self.use_prefix_l1: bool = False
        self.use_prefix_deep: bool = False
        self.use_file_root: bool = True
        self.node_name_strategy: str = "first"
        self.multi_sub_delimiter: str = ";"
 
        # Global Data Separator
        self.data_delimiter: str = ";"
 
        # NEW: Exclude title column value from node data elements
        self.exclude_title_from_data: bool = True
 
    @staticmethod
    def _clean_headers(headers) -> list[str]:
        res = []
        counts = {}
        for i, h in enumerate(headers):
            if pd.isna(h) or str(h).strip() == "":
                clean_name = f"Col{i+1}"
            else:
                s = str(h).strip()
                s = re.sub(r'[^0-9A-Za-zА-Яа-яЁёІіЇїЄєҐґ _\-\.]', '', s)
                s = re.sub(r'\s+', '_', s).strip('_')
                clean_name = s if s else f"Col{i+1}"
 
            if clean_name in counts:
                counts[clean_name] += 1
                final_name = f"{clean_name}_{counts[clean_name]}"
            else:
                counts[clean_name] = 1
                final_name = clean_name
            res.append(final_name)
        return res
 
    def _find_header_row(self, df_preview: pd.DataFrame) -> int:
        best_idx = 0
        max_valid_cols = -1
        limit = min(20, len(df_preview))
        for i in range(limit):
            row = df_preview.iloc[i]
            valid_cols = 0
            for val in row:
                if pd.notna(val) and str(val).strip() != "":
                    valid_cols += 1
            if valid_cols > max_valid_cols:
                max_valid_cols = valid_cols
                best_idx = i
        return best_idx
 
    def _reset_state(self):
        self.hierarchy_configs = {}
        self.manual_grouping_rules = []
        self.grouping_rules = []
        self.column_type_preferences = {}
        self.sheet_title_cols = {}
 
    # ------------------------------------------------------------------
    # NEW: Inject hyperlinks from openpyxl into a pandas DataFrame so that
    #      both =HYPERLINK() formula cells and Ctrl+K hyperlink cells are
    #      encoded as formula strings that _extract_url_from_formula() can
    #      parse later.
    # ------------------------------------------------------------------
    @staticmethod
    def _inject_hyperlinks_from_ws(df: pd.DataFrame, ws_opx, header_idx: int) -> pd.DataFrame:
        """
        Walks the openpyxl worksheet and, for every data cell that carries a
        hyperlink (Ctrl+K) or contains an =HYPERLINK() formula, replaces the
        corresponding DataFrame value with '=HYPERLINK("url","text")' so the
        existing _extract_url_from_formula logic handles it uniformly.
 
        Parameters
        ----------
        df          : DataFrame already read by pandas (header already applied,
                      rows indexed 0 … len-1 matching the worksheet data rows).
        ws_opx      : openpyxl Worksheet object (loaded without data_only).
        header_idx  : 0-based index of the header row inside the sheet as
                      detected by _find_header_row().
        """
        col_names = list(df.columns)
        # openpyxl rows are 1-based; the first data row in openpyxl is:
        #   header_idx (0-based) + 1 (to make it 1-based) + 1 (skip header row itself)
        data_start_1idx = header_idx + 2
 
        for opx_row in ws_opx.iter_rows(min_row=data_start_1idx):
            df_row_idx = opx_row[0].row - data_start_1idx  # 0-based df row
            if df_row_idx < 0 or df_row_idx >= len(df):
                continue
 
            for cell in opx_row:
                col_0idx = cell.column - 1  # 0-based column index
                if col_0idx >= len(col_names):
                    continue
 
                # ── Case 1: cell contains a =HYPERLINK() formula ──────────────
                # When openpyxl is opened without data_only=True, formula cells
                # expose the raw formula string as cell.value.
                cell_val = cell.value
                if isinstance(cell_val, str) and re.match(r'\s*=HYPERLINK\s*\(', cell_val, re.IGNORECASE):
                    # Inject the formula string directly; pandas will have stored
                    # only the cached result, so we overwrite with the formula.
                    df.iat[df_row_idx, col_0idx] = cell_val
                    continue
 
                # ── Case 2: Ctrl+K / worksheet-level hyperlink ────────────────
                if cell.hyperlink:
                    hl = cell.hyperlink
                    url = ""
                    if hasattr(hl, 'target') and hl.target:
                        url = hl.target
                    elif isinstance(hl, str):
                        url = hl
                    if url and url.startswith(('http://', 'https://', 'ftp://')):
                        # The current df value is the display text
                        current = df.iat[df_row_idx, col_0idx]
                        display = str(current).strip() if pd.notna(current) else ""
                        # Sanitise to avoid breaking the simple regex parser
                        display_safe = display.replace('"', "'")
                        url_safe = url.replace('"', '%22')
                        df.iat[df_row_idx, col_0idx] = f'=HYPERLINK("{url_safe}","{display_safe}")'
 
        return df
 
    def load_xlsx(self, file_path: str) -> tuple[bool, str, list[str]]:
        try:
            import openpyxl  # already a pandas dependency, always available
            # Load workbook WITHOUT data_only so formula strings are visible
            wb = openpyxl.load_workbook(file_path, data_only=False)
 
            xl = pd.ExcelFile(file_path, engine="openpyxl")
            self.sheets_data, self.loaded_sheets = {}, []
 
            for sheet in xl.sheet_names:
                df_raw = pd.read_excel(xl, sheet_name=sheet, header=None, nrows=25)
                if df_raw.empty:
                    continue
                header_idx = self._find_header_row(df_raw)
                df = pd.read_excel(xl, sheet_name=sheet, header=header_idx)
                df.columns = self._clean_headers(df.columns)
 
                # ── NEW: inject hyperlink URLs ────────────────────────────────
                if sheet in wb.sheetnames:
                    df = self._inject_hyperlinks_from_ws(df, wb[sheet], header_idx)
                # ─────────────────────────────────────────────────────────────
 
                df = df.replace('', pd.NA).dropna(how='all')
                if not df.empty:
                    self.sheets_data[sheet] = df
                    self.loaded_sheets.append(sheet)
 
            if not self.loaded_sheets:
                return False, "No valid sheets found.", []
 
            raw_name = os.path.splitext(os.path.basename(file_path))[0]
            self.root_name = re.sub(r'[\\/*?:"<>|]', "", raw_name)
 
            self.selected_sheets = set(s for s in self.loaded_sheets if s != "structure")
            self._reset_state()
            return True, f"Loaded {len(self.loaded_sheets)} sheets.", [s for s in self.loaded_sheets if s != "structure"]
        except Exception as e:
            return False, f"Excel Error: {str(e)}", []
 
    def load_gsheet(self, url: str, creds_json: str) -> tuple[bool, str, list[str]]:
        if not HAS_SHEETS:
            return False, "Error: 'gspread' library not installed.", []
        try:
            if not url or not creds_json:
                return False, "Please provide URL and JSON.", []
            creds_dict = json.loads(creds_json)
            scope = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
            creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
            client = gspread.authorize(creds)
            sheet = client.open_by_url(url)
            self.root_name = re.sub(r'[\\/*?:"<>|]', "", sheet.title)
            self.sheets_data, self.loaded_sheets = {}, []
            for ws in sheet.worksheets():
                # ── NEW: request formula strings so =HYPERLINK() is preserved ──
                try:
                    data = ws.get(value_render_option='FORMULA')
                except Exception:
                    data = ws.get_all_values()
                # ──────────────────────────────────────────────────────────────
                if not data:
                    continue
                df_raw = pd.DataFrame(data)
                header_idx = self._find_header_row(df_raw)
                if header_idx + 1 < len(data):
                    headers = data[header_idx]
                    rows = data[header_idx+1:]
                    clean_headers = self._clean_headers(headers)
                    df = pd.DataFrame(rows, columns=clean_headers)
                    df = df.replace('', pd.NA).dropna(how='all')
                    if not df.empty:
                        self.sheets_data[ws.title] = df
                        self.loaded_sheets.append(ws.title)
            if not self.loaded_sheets:
                return False, "No valid data.", []
            self.selected_sheets = set(s for s in self.loaded_sheets if s != "structure")
            self._reset_state()
            return True, f"Loaded {len(self.loaded_sheets)} sheets.", [s for s in self.loaded_sheets if s != "structure"]
        except Exception as e:
            return False, f"Error: {str(e)}", []
 
    def get_eligible_entities_for_types(self) -> list[str]:
        eligible = set()
        eligible.update(self.selected_sheets)
        eligible.update(self.hierarchy_configs.keys())
        return sorted(list(eligible))
 
    @staticmethod
    def _extract_url_from_formula(cell_value: str) -> tuple[str, str, bool]:
        """Returns (display_text, url, has_link).
        Handles:
          - =HYPERLINK("url","text") formulas
          - Plain URLs (http/https)
          - Plain text (no link)
        """
        if not isinstance(cell_value, str):
            return (str(cell_value), "", False)
        # HYPERLINK formula
        m = re.search(r'=HYPERLINK\s*\(\s*"([^"]+)"\s*,\s*"([^"]*)"\s*\)', cell_value, re.IGNORECASE)
        if m:
            return (m.group(2) or m.group(1), m.group(1), True)
        # Plain URL
        if re.match(r'https?://[^\s]+', cell_value.strip()):
            return ("", cell_value.strip(), True)
        return (cell_value, "", False)
 
    @staticmethod
    def _is_image(url: str) -> bool:
        if not url or not isinstance(url, str): return False
        url_path = url.split('?')[0].lower().strip()
        return any(url_path.endswith(e) for e in [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".svg", ".webp", ".jfif"])
 
    @staticmethod
    def _is_youtube(url: str) -> bool:
        if not url or not isinstance(url, str): return False
        u = url.lower().strip()
        return "youtube.com" in u or "youtu.be" in u
 
    def _detect_type(self, url_or_text: str) -> str:
        if self._is_image(url_or_text): return "image"
        if self._is_youtube(url_or_text): return "youtube"
        return "link"
 
    def _make_data_element(self, col: str, v: str, sheet: str) -> dict:
        """Convert a single value string into a data element dict with proper type detection."""
        text, url, has_link = self._extract_url_from_formula(v)
        pref = self.column_type_preferences.get(sheet, {}).get(col, "auto")
 
        if has_link:
            if pref == "auto":
                ltype = self._detect_type(url)
            else:
                ltype = pref
            return {"tclass": str(col), "link": url, "type": ltype, "content": text or ""}
        else:
            ltype = "text" if pref == "auto" else pref
            return {"tclass": str(col), "link": "", "type": ltype, "content": v}
 
    @staticmethod
    def _format_value(val) -> str:
        """Convert a cell value to a clean string.
        Timestamps / date objects → DD.MM.YYYY.
        ISO-like strings (2023-01-15 or 2023-01-15 00:00:00) → DD.MM.YYYY.
        Everything else → plain str().
        """
        if hasattr(val, 'strftime'):
            return val.strftime('%d.%m.%Y')
        s = str(val).strip()
        ts_match = re.match(r'^(\d{4})-(\d{2})-(\d{2})(?:\s+\d{2}:\d{2}:\d{2}(?:\.\d+)?)?$', s)
        if ts_match:
            return f"{ts_match.group(3)}.{ts_match.group(2)}.{ts_match.group(1)}"
        return s
 
    def _row_to_elements(self, row: pd.Series, cols: list[str], sheet: str) -> list[dict]:
        result = []
        # NEW: determine which column (if any) is the title and should be excluded
        excluded_col = self.sheet_title_cols.get(sheet) if self.exclude_title_from_data else None
 
        for col in cols:
            # NEW: skip the title column so its value isn't duplicated as a data element
            if excluded_col and col == excluded_col:
                continue
 
            val = row.get(col)
            if pd.notna(val) and str(val).strip():
                v_str = self._format_value(val)
                values = [v_str]
                if self.data_delimiter and self.data_delimiter != "No Separation":
                    values = [x.strip() for x in v_str.split(self.data_delimiter) if x.strip()]
 
                for v in values:
                    result.append(self._make_data_element(col, v, sheet))
        return result
 
    def _create_edge(self, parent: str, child: str) -> GraphEdge:
        return GraphEdge(parent, child) if self.edge_direction == "left" else GraphEdge(child, parent)
 
    def _get_or_create_row_node(self, sheet_name: str, row: pd.Series, cols: list[str]) -> GraphNode:
        node_id = f"{sheet_name}_{row.name}"
        if node_id in self.node_map: return self.node_map[node_id]
        title = None
        configured_col = self.sheet_title_cols.get(sheet_name)
        if configured_col and configured_col in cols:
            val = row.get(configured_col)
            if pd.notna(val) and str(val).strip(): title = str(val).strip()
        if not title and self.node_name_strategy == "Name":
            name_col = next((c for c in cols if c.lower() == "name"), None)
            if name_col:
                val = row.get(name_col)
                if pd.notna(val) and str(val).strip(): title = str(val).strip()
        if not title and len(cols) > 0:
            val = row.iloc[0]
            if pd.notna(val) and str(val).strip(): title = str(val).strip()
        if not title:
            for c in cols:
                val = row.get(c)
                if pd.notna(val) and str(val).strip() and not str(val).isnumeric():
                    title = str(val).strip()
                    break
        if not title: title = f"{sheet_name}_{row.name}"
        if title in self.node_map: return self.node_map[title]
        node = GraphNode(title, self._row_to_elements(row, cols, sheet_name), sheet_name)
        self.nodes.append(node)
        self.node_map[title] = node
        self.node_map[node_id] = node
        return node
 
    def _lookup_in_structure(self, val: str) -> tuple[t.Optional[pd.Series], t.Optional[pd.DataFrame], t.Optional[str]]:
        struct_df = self.sheets_data.get("structure")
        if struct_df is not None and len(struct_df.columns) > 0:
            title_col = struct_df.columns[0]
            matches = struct_df[struct_df[title_col].astype(str).str.strip() == str(val).strip()]
            if not matches.empty:
                return matches.iloc[0], struct_df, "structure"
        return None, None, None
 
    def _enrich_from_row(self, match_row: pd.Series, lookup_df: pd.DataFrame, source_sheet: str) -> list[dict]:
        data_elements = []
        lookup_cols = list(lookup_df.columns)
        for lookup_col in lookup_cols:
            elem_val = match_row.get(lookup_col)
            if pd.notna(elem_val) and str(elem_val).strip():
                v_str = self._format_value(elem_val)
                values = [v_str]
                if self.data_delimiter and self.data_delimiter != "No Separation":
                    values = [x.strip() for x in v_str.split(self.data_delimiter) if x.strip()]
 
                for v in values:
                    data_elements.append(self._make_data_element(lookup_col, v, source_sheet))
        return data_elements
 
    def _enrich_phantom_data(self, col: str, val: str, ctx: tuple) -> list[dict]:
        match_row, lookup_df, source_sheet = self._lookup_in_structure(val)
 
        if match_row is None and col in self.sheets_data:
            lookup_df = self.sheets_data[col]
            if len(lookup_df.columns) > 0:
                title_col = lookup_df.columns[0]
                matches = lookup_df[lookup_df[title_col].astype(str).str.strip() == str(val).strip()]
                if not matches.empty:
                    match_row = matches.iloc[0]
                    source_sheet = col
        if match_row is not None and lookup_df is not None:
            return self._enrich_from_row(match_row, lookup_df, source_sheet)
        return []
 
    def _enrich_node_from_structure(self, node_name: str) -> list[dict]:
        match_row, lookup_df, source_sheet = self._lookup_in_structure(node_name)
        if match_row is not None and lookup_df is not None:
            return self._enrich_from_row(match_row, lookup_df, source_sheet)
        return []
 
    def _get_phantom(self, ctx: tuple, col: str, val: str) -> str:
        key = (ctx, col, val)
        if key in self.phantom_nodes: return self.phantom_nodes[key]
 
        is_level_1 = (len(ctx) == 0)
        use_prefix = self.use_prefix_l1 if is_level_1 else self.use_prefix_deep
 
        name = f"{col}: {val}" if use_prefix else f"{val}"
        data_elements = self._enrich_phantom_data(col, val, ctx)
        if name in self.node_map:
            existing = self.node_map[name]
            if not existing.data_elements and data_elements:
                existing.data_elements = data_elements
            self.phantom_nodes[key] = name
            return name
        node = GraphNode(name, data_elements, col)
        self.nodes.append(node)
        self.node_map[name] = node
        self.phantom_nodes[key] = name
        return node.name
 
    def _split_multi_value(self, value: str, delimiter: str = None) -> list[str]:
        if not delimiter or delimiter == "No Separation": return [value]
        return [v.strip() for v in str(value).split(delimiter) if v.strip()]
 
    def _apply_mode4_rule(self, rule: dict):
        if "pairs" in rule:
            group_name = rule.get("name", "Group")
            delimiter = rule.get("delimiter", self.multi_sub_delimiter)
            fields = []
            for s, f in rule["pairs"]:
                if s in self.sheets_data and s in self.selected_sheets:
                    df = self.sheets_data[s]
                    if f in df.columns:
                        vals = set(str(v).strip() for v in df[f].dropna() if str(v).strip())
                        fields.append((s, f, vals))
        else:
            return
 
        if not fields: return
 
        for s, _, _ in fields: self.sheets_in_grouping.add(s)
 
        value_to_sf = {}
        for s, f, _ in fields:
            df = self.sheets_data[s]
            for val in df[f].dropna():
                v_str = str(val).strip()
                if not v_str: continue
                sub_values = self._split_multi_value(v_str, delimiter)
                for sub_val in sub_values: value_to_sf.setdefault(sub_val, []).append((s, f, val))
 
        for value, sheet_field_originals in value_to_sf.items():
            v_key = f"M4_{value}"
            data_elements = self._enrich_phantom_data(group_name, value, ())
            if not data_elements:
                data_elements = self._enrich_node_from_structure(value)
            if v_key not in self.node_map:
                v_node = GraphNode(value, data_elements, group_name, is_intermediate=True)
                self.nodes.append(v_node)
                self.node_map[v_key] = v_node
                self.node_map[value] = v_node
            elif data_elements and not self.node_map[v_key].data_elements:
                self.node_map[v_key].data_elements = data_elements
            self.edges.append(self._create_edge(self.root_name, value))
 
            processed_combos = set()
            for s, f, orig_val in sheet_field_originals:
                combo_key = (s, f, orig_val)
                if combo_key in processed_combos: continue
                processed_combos.add(combo_key)
 
                root_node_name = s
                if root_node_name not in self.node_map:
                    sheet_data = self._enrich_node_from_structure(root_node_name)
                    root_node = GraphNode(root_node_name, sheet_data, s)
                    self.nodes.append(root_node)
                    self.node_map[root_node_name] = root_node
                self.edges.append(self._create_edge(value, root_node_name))
 
                matches = self.sheets_data[s][self.sheets_data[s][f].astype(str) == str(orig_val)]
                cfg = self.hierarchy_configs.get(s, [])
                if cfg:
                    self._build_hier(matches.reset_index(drop=True), cfg, root_node_name, s, 0)
                else:
                    for _, row in matches.iterrows():
                        node = self._get_or_create_row_node(s, row, list(self.sheets_data[s].columns))
                        self.edges.append(self._create_edge(root_node_name, node.name))
 
        if self.show_empty_nodes:
            empty_sheet_rows = []
            for s, f, _ in fields:
                df_s = self.sheets_data[s]
                empty_mask = df_s[f].isna() | (df_s[f].astype(str).str.strip() == "")
                empty_rows = df_s[empty_mask]
                if not empty_rows.empty: empty_sheet_rows.append((s, f, empty_rows))
            if empty_sheet_rows:
                empty_key = "(empty)"
                empty_data = []
                if empty_key not in self.node_map:
                    empty_node = GraphNode("empty", empty_data, group_name, is_intermediate=True)
                    self.nodes.append(empty_node)
                    self.node_map[empty_key] = empty_node
                    self.node_map["empty"] = empty_node
                self.edges.append(self._create_edge(self.root_name, "empty"))
                for s, f, empty_rows in empty_sheet_rows:
                    if s not in self.node_map:
                        sheet_data = self._enrich_node_from_structure(s)
                        s_node = GraphNode(s, sheet_data, s)
                        self.nodes.append(s_node)
                        self.node_map[s] = s_node
                    self.edges.append(self._create_edge("empty", s))
                    cfg = self.hierarchy_configs.get(s, [])
                    if cfg:
                        self._build_hier(empty_rows.reset_index(drop=True), cfg, s, s, 0)
                    else:
                        for _, row in empty_rows.iterrows():
                            node = self._get_or_create_row_node(s, row, list(self.sheets_data[s].columns))
                            self.edges.append(self._create_edge(s, node.name))
 
    def apply_groupings(self):
        if not self.grouping_enabled and not self.manual_grouping_rules: return
        self.sheets_in_grouping = set()
        for rule in self.grouping_rules + self.manual_grouping_rules: self._apply_mode4_rule(rule)
 
    def _build_hier(self, df: pd.DataFrame, cfg: list[dict], parent: str, sheet: str, lvl: int = 0, current_filter: dict = None):
        if current_filter is None: current_filter = {}
        if lvl >= len(cfg):
            filtered_rows = df.copy()
            for col, val in current_filter.items():
                if val is None:
                    filtered_rows = filtered_rows[filtered_rows[col].isna() | (filtered_rows[col].astype(str).str.strip() == "")]
                else:
                    filtered_rows = filtered_rows[filtered_rows[col].astype(str) == str(val)]
            for _, row in filtered_rows.iterrows():
                node = self._get_or_create_row_node(sheet, row, list(df.columns))
                self.edges.append(self._create_edge(parent, node.name))
            return
 
        level_conf = cfg[lvl]
        if isinstance(level_conf, dict):
            col = level_conf.get('col')
            level_delim = level_conf.get('delim', ';')
        else:
            col = str(level_conf)
            level_delim = ';'
 
        if col not in df.columns: return
 
        curr = df.copy()
        for col_name, val in current_filter.items():
            if val is None:
                curr = curr[curr[col_name].isna() | (curr[col_name].astype(str).str.strip() == "")]
            else:
                curr = curr[curr[col_name].astype(str) == str(val)]
 
        processed_vals = set()
        for index, row in curr.iterrows():
            val = row.get(col)
            if pd.isna(val) or not str(val).strip(): continue
            val_str = str(val).strip()
 
            if val_str in processed_vals: continue
            processed_vals.add(val_str)
 
            sub_values = self._split_multi_value(val_str, level_delim)
 
            for sub_val in sub_values:
                phantom_name = self._get_phantom(tuple(sorted([(c, str(v)) for c, v in current_filter.items()])), col, sub_val)
                edge_exists = any(e.node1 == parent and e.node2 == phantom_name for e in self.edges)
                if not edge_exists: self.edges.append(self._create_edge(parent, phantom_name))
                nf = dict(current_filter)
                nf[col] = val_str
                self._build_hier(df, cfg, phantom_name, sheet, lvl+1, nf)
 
        if self.show_empty_nodes:
            empty_mask = curr[col].isna() | (curr[col].astype(str).str.strip() == "")
            if empty_mask.any():
                empty_val = "(empty)"
                empty_name = self._get_phantom(tuple(sorted([(c, str(v)) for c, v in current_filter.items()])), col, empty_val)
                edge_exists = any(e.node1 == parent and e.node2 == empty_name for e in self.edges)
                if not edge_exists: self.edges.append(self._create_edge(parent, empty_name))
                nf = dict(current_filter)
                nf[col] = None
                self._build_hier(df, cfg, empty_name, sheet, lvl+1, nf)
 
    def generate(self, mode_val: str) -> tuple[str, int, int]:
        if not self.selected_sheets and not self.loaded_sheets: raise RuntimeError("No sheets loaded")
        if mode_val == "Typical (Simple)":
            self.edge_direction = "right"
            self.node_name_strategy = "first"
            if not self.selected_sheets: self.selected_sheets = set(s for s in self.loaded_sheets if s != "structure")
 
        self.nodes, self.edges, self.node_map, self.phantom_nodes = [], [], {}, {}
 
        if self.use_file_root:
            root_data = self._enrich_node_from_structure(self.root_name)
            root = GraphNode(self.root_name, root_data)
            self.nodes.append(root)
            self.node_map[self.root_name] = root
 
        self.apply_groupings()
 
        for name in sorted(self.selected_sheets):
            if name not in self.sheets_data or name in self.sheets_in_grouping: continue
            if name not in self.node_map:
                sheet_data = self._enrich_node_from_structure(name)
                n = GraphNode(name, sheet_data, name)
                self.nodes.append(n)
                self.node_map[name] = n
            else:
                existing = self.node_map[name]
                if not existing.data_elements:
                    existing.data_elements = self._enrich_node_from_structure(name)
            if self.use_file_root:
                self.edges.append(self._create_edge(self.root_name, name))
 
            df = self.sheets_data[name]
            cfg = self.hierarchy_configs.get(name, [])
            if cfg: self._build_hier(df, cfg, name, name, 0)
            else:
                for _, row in df.iterrows():
                    node = self._get_or_create_row_node(name, row, list(df.columns))
                    self.edges.append(self._create_edge(name, node.name))
 
        connected = set()
        for e in self.edges:
            connected.add(e.node1)
            connected.add(e.node2)
        self.nodes = [n for n in self.nodes if n.name in connected]
 
        graph = ET.Element("Graph")
        graph.set("vspacing", "30")
        graph.set("hspacing", "200")
        graph.set("padding", "40")
        graph.set("guid", uuid.uuid4().hex[:12].upper())
        nodes_el = ET.SubElement(graph, "Nodes")
        x, y = 0, 0
        for i, n in enumerate(self.nodes):
            n.x_pos, n.y_pos = x, y
            el = ET.SubElement(nodes_el, "Node")
            el.set("guid", n.guid)
            el.set("nodeName", n.name)
            el.set("nclass", n.node_type)
            el.set("shape", n.shape)
            el.set("xPos", str(n.x_pos))
            el.set("yPos", str(n.y_pos))
            el.set("color", n.color)
            for d in n.data_elements:
                de = ET.SubElement(el, "data")
                de.set("tclass", d["tclass"])
                de.set("link", d["link"])
                de.set("type", d["type"])
                de.text = d["content"]
            x += 300
            if (i+1) % 5 == 0: x, y = 0, y+150
        edges_el = ET.SubElement(graph, "Edges")
        for e in self.edges:
            ee = ET.SubElement(edges_el, "Edge")
            ee.set("guid", e.guid)
            ee.set("group", e.group)
            ee.set("node1", e.node1)
            ee.set("node2", e.node2)
            ee.set("bilateral", "true" if e.bilateral else "")
        ET.SubElement(graph, "Groups")
 
        safe_name = re.sub(r'[\\/*?:"<>|]', "", self.root_name) or "Graph"
        out_file = f"{safe_name}.xml"
 
        xml_str = minidom.parseString(ET.tostring(graph, encoding="unicode")).toprettyxml(indent="  ")
        with open(out_file, "w", encoding="utf-8") as f: f.write(xml_str)
        return out_file, len(self.nodes), len(self.edges)
 
 
# --- [UI STYLE] ---
css = """
:root { --primary: #2563eb; --bg: #f3f4f6; --card-bg: #ffffff; }
body, .gradio-container { background: var(--bg); font-family: 'Inter', sans-serif; }
.step-card { background: var(--card-bg); border: 1px solid #e5e7eb; border-radius: 12px; padding: 20px; margin-bottom: 16px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }
.step-header { font-size: 16px; font-weight: 700; color: #1f2937; margin-bottom: 16px; padding-bottom: 8px; border-bottom: 2px solid #f3f4f6; text-transform: uppercase; letter-spacing: 0.05em; }
.clean-row { gap: 16px; align-items: flex-end; margin-bottom: 8px; }
.btn-main { background-color: #2563eb !important; color: white !important; font-size: 15px; font-weight: 600; height: 48px; }
.btn-action { background-color: #ffffff !important; color: #374151 !important; border: 1px solid #d1d5db !important; height: 40px; font-size: 13px; }
.btn-del { background-color: #fee2e2 !important; color: #b91c1c !important; border: 1px solid #fca5a5 !important; height: 40px; font-size: 13px; }
input, select, textarea { font-size: 14px !important; border-radius: 6px !important; }
textarea { height: 80px !important; }
.tpl-bar { background: linear-gradient(135deg, #eff6ff 0%, #f5f3ff 100%) !important; border-color: #c7d2fe !important; }
"""
 
conv = Converter()

# ─────────────────────────────────────────────────────────────
# TEMPLATE PERSISTENCE HELPERS
# ─────────────────────────────────────────────────────────────
TEMPLATES_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates.json")


def _tmpl_load_all() -> dict:
    if os.path.exists(TEMPLATES_FILE):
        try:
            with open(TEMPLATES_FILE, "r", encoding="utf-8") as fh:
                return json.load(fh)
        except Exception:
            return {}
    return {}


def _tmpl_save_all(data: dict) -> None:
    with open(TEMPLATES_FILE, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)


def _tmpl_snapshot_structure() -> dict:
    """Capture {sheet: [col, ...]} for the currently-loaded file."""
    return {s: list(df.columns) for s, df in conv.sheets_data.items()}


def _tmpl_collect_config() -> dict:
    """Collect all Converter settings into a serialisable dict."""
    return {
        "column_type_preferences": conv.column_type_preferences,
        "hierarchy_configs": {
            s: [
                {"col": lv["col"], "delim": lv.get("delim")} if isinstance(lv, dict)
                else {"col": str(lv), "delim": None}
                for lv in levels
            ]
            for s, levels in conv.hierarchy_configs.items()
        },
        "sheet_title_cols": conv.sheet_title_cols,
        "manual_grouping_rules": [
            {
                "name": r.get("name", "Group"),
                "pairs": [[e, c] for e, c in r.get("pairs", [])],
                "delimiter": r.get("delimiter"),
            }
            for r in conv.manual_grouping_rules
        ],
        "selected_sheets": sorted(list(conv.selected_sheets)),
        "data_delimiter": conv.data_delimiter,
        "edge_direction": conv.edge_direction,
        "use_prefix_l1": conv.use_prefix_l1,
        "use_prefix_deep": conv.use_prefix_deep,
        "use_file_root": conv.use_file_root,
        "exclude_title_from_data": conv.exclude_title_from_data,
    }


def _auto_save():
    """Silently persist current config to templates.json under '_autosave' key."""
    if not conv.loaded_sheets:
        return
    try:
        all_tpls = _tmpl_load_all()
        user_tpls = all_tpls.get("_autosave", {})
        name = conv.root_name or "session"
        user_tpls[name] = {
            "created": datetime.datetime.now().isoformat(),
            "structural_metadata": _tmpl_snapshot_structure(),
            "config": _tmpl_collect_config(),
        }
        all_tpls["_autosave"] = user_tpls
        _tmpl_save_all(all_tpls)
    except Exception:
        pass  # Never disrupt the user on auto-save failure


def _tmpl_apply(template: dict) -> tuple:
    """
    Apply a saved template to conv, respecting the current file structure.
    Returns (applied_msgs, skipped_msgs, match_type).
    match_type: "full" | "partial" | "none"
    """
    config = template["config"]
    current_meta = _tmpl_snapshot_structure()

    applied: list = []
    skipped: list = []
    total_sc = 0
    applicable_sc = 0

    def _check(sheet: str, col: str) -> bool:
        nonlocal total_sc, applicable_sc
        total_sc += 1
        ok = sheet in current_meta and col in current_meta[sheet]
        if ok:
            applicable_sc += 1
        return ok

    # Pre-scan all sheet+column-specific settings for threshold calculation
    for sheet, col_prefs in config.get("column_type_preferences", {}).items():
        for col in col_prefs:
            _check(sheet, col)
    for sheet, levels in config.get("hierarchy_configs", {}).items():
        for lv in levels:
            _check(sheet, lv["col"] if isinstance(lv, dict) else str(lv))
    for sheet, col in config.get("sheet_title_cols", {}).items():
        _check(sheet, col)
    for rule in config.get("manual_grouping_rules", []):
        for pair in rule.get("pairs", []):
            _check(pair[0], pair[1])

    if total_sc == 0:
        match_type = "full"
    elif applicable_sc == total_sc:
        match_type = "full"
    elif applicable_sc > total_sc / 2:
        match_type = "partial"
    else:
        match_type = "none"

    if match_type == "none" and total_sc > 0:
        return applied, skipped, match_type

    # ── Global settings (always applied) ──────────────────────
    for key in ("data_delimiter", "edge_direction", "use_prefix_l1", "use_prefix_deep",
                "use_file_root", "exclude_title_from_data"):
        if key in config:
            setattr(conv, key, config[key])
            applied.append(f"Setting `{key}` = `{config[key]}`")

    # ── selected_sheets ───────────────────────────────────────
    saved_sel = config.get("selected_sheets", [])
    new_sel = [s for s in saved_sel if s in current_meta]
    for s in saved_sel:
        if s not in current_meta:
            skipped.append(f"Sheet '{s}' not found — removed from selection")
    if new_sel:
        conv.selected_sheets = set(new_sel)
        applied.append(f"Selected sheets: {', '.join(new_sel)}")

    # ── column_type_preferences ───────────────────────────────
    new_ctp: dict = {}
    for sheet, col_prefs in config.get("column_type_preferences", {}).items():
        if sheet not in current_meta:
            for col in col_prefs:
                skipped.append(f"Sheet '{sheet}' not found — type pref for '{col}' skipped")
            continue
        sp: dict = {}
        for col, pref in col_prefs.items():
            if col in current_meta[sheet]:
                sp[col] = pref
                applied.append(f"Column type `{sheet}.{col}` → `{pref}`")
            else:
                skipped.append(f"Column '{col}' not in '{sheet}' — type pref skipped")
        if sp:
            new_ctp[sheet] = sp
    conv.column_type_preferences = new_ctp

    # ── hierarchy_configs ─────────────────────────────────────
    new_hc: dict = {}
    for sheet, levels in config.get("hierarchy_configs", {}).items():
        if sheet not in current_meta:
            skipped.append(f"Sheet '{sheet}' not found — {len(levels)} hierarchy level(s) skipped")
            continue
        valid: list = []
        for lv in levels:
            col = lv["col"] if isinstance(lv, dict) else str(lv)
            if col in current_meta[sheet]:
                valid.append(lv)
                applied.append(f"Hierarchy `{sheet}.{col}`")
            else:
                skipped.append(f"Column '{col}' not in '{sheet}' — hierarchy level skipped")
        if valid:
            new_hc[sheet] = valid
    conv.hierarchy_configs = new_hc

    # ── sheet_title_cols ──────────────────────────────────────
    new_stc: dict = {}
    for sheet, col in config.get("sheet_title_cols", {}).items():
        if sheet not in current_meta:
            skipped.append(f"Sheet '{sheet}' not found — title column skipped")
        elif col not in current_meta[sheet]:
            skipped.append(f"Column '{col}' not in '{sheet}' — title column skipped")
        else:
            new_stc[sheet] = col
            applied.append(f"Title column `{sheet}` → `{col}`")
    conv.sheet_title_cols = new_stc

    # ── manual_grouping_rules ─────────────────────────────────
    new_rules: list = []
    for rule in config.get("manual_grouping_rules", []):
        valid_pairs: list = []
        for pair in rule.get("pairs", []):
            entity, col = pair[0], pair[1]
            if entity not in current_meta:
                skipped.append(f"Sheet '{entity}' not found — grouping '{rule.get('name', '')}' pair skipped")
            elif col not in current_meta[entity]:
                skipped.append(f"Column '{col}' not in '{entity}' — grouping '{rule.get('name', '')}' pair skipped")
            else:
                valid_pairs.append([entity, col])
                applied.append(f"Grouping '{rule.get('name', 'Group')}': `{entity}.{col}`")
        if valid_pairs:
            new_rules.append({**rule, "pairs": valid_pairs})
    conv.manual_grouping_rules = new_rules

    return applied, skipped, match_type


with gr.Blocks(css=css, title="Graph Converter") as app:
    # ─── TEMPLATE MANAGEMENT BAR ───────────────────────────────
    with gr.Group(elem_classes=["step-card", "tpl-bar"]):
        gr.HTML('<div class="step-header" style="color:#4f46e5; border-color:#c7d2fe;">Template Management</div>')
        with gr.Row(elem_classes=["clean-row"]):
            tpl_login = gr.Textbox(label="Login", placeholder="Enter your login identifier", scale=3)
            tpl_save_btn = gr.Button("💾 Save Template", elem_classes=["btn-action"], scale=1)
            tpl_load_btn = gr.Button("📂 Load Template", elem_classes=["btn-action"], scale=1)

        with gr.Column(visible=False) as tpl_save_panel:
            with gr.Row(elem_classes=["clean-row"]):
                tpl_name_input = gr.Textbox(label="Template Name", placeholder="Leave empty for auto-generated name", scale=3)
                tpl_overwrite_cb = gr.Checkbox(label="Overwrite if exists", value=False, visible=False, scale=1)
                tpl_confirm_save_btn = gr.Button("✅ Confirm Save", variant="primary", elem_classes=["btn-main"], scale=1)

        with gr.Column(visible=False) as tpl_load_panel:
            with gr.Row(elem_classes=["clean-row"]):
                tpl_list = gr.Dropdown(choices=[], label="Select Template", scale=3)
                tpl_confirm_load_btn = gr.Button("✅ Apply Template", variant="primary", elem_classes=["btn-main"], scale=1)

        tpl_status = gr.Markdown(value="")
        with gr.Accordion("Applied / Skipped Settings", open=True, visible=False) as tpl_summary_accordion:
            tpl_summary = gr.Markdown()

    # --- STEP 1 ---
    with gr.Group(elem_classes=["step-card"]):
        gr.HTML('<div class="step-header">Step 1: Setup & Source</div>')
        with gr.Tabs():
            with gr.Tab("📂 Upload Excel"):
                with gr.Row(elem_classes=["clean-row"]):
                    up_file = gr.File(label="Excel File (.xlsx)", file_count="single", file_types=[".xlsx"])
                    load_file_btn = gr.Button("Load Excel", variant="primary", elem_classes=["btn-main"])
            with gr.Tab("☁️ Google Sheets"):
                gr.Markdown("Enter URL and Service Account JSON (saved in your browser).")
                gs_url = gr.Textbox(label="Google Sheet URL")
                gs_json = gr.Textbox(label="Service Account JSON", lines=3, elem_id="gs_creds")
                load_gs_btn = gr.Button("Load Google Sheet", variant="primary", elem_classes=["btn-main"])
        gr.HTML('<div style="height:16px"></div>')
        with gr.Row(elem_classes=["clean-row"]):
            mode_toggle = gr.Radio(choices=["Typical (Simple)", "Advanced"], value="Typical (Simple)", label="Mode", info="Advanced mode unlocks grouping and custom types.")
            with gr.Column():
                use_file_root_cb = gr.Checkbox(label="Use File Name as Root Node", value=True)
                # NEW: Exclude title column from node data elements
                exclude_title_cb = gr.Checkbox(
                    label="Exclude Title Column from Node Data",
                    value=True,
                    info="When on, the column used as the node title is not duplicated as a data attribute."
                )
            global_data_delim = gr.Dropdown(choices=DELIMITER_CHOICES, value=";", label="Data Field Separator")
        status = gr.Markdown(value="Ready to load data.")
 
    with gr.Row():
        # --- STEP 2 ---
        with gr.Column(scale=3):
            with gr.Group(elem_classes=["step-card"]):
                gr.HTML('<div class="step-header" style="color:#2563eb; border-color:#bfdbfe;">Step 2: Hierarchy</div>')
                with gr.Row(elem_classes=["clean-row"]):
                    hier_sheet = gr.Dropdown(choices=[], label="Select Entity (Sheet)", scale=2)
                    hier_add_col = gr.Dropdown(choices=[], label="Select Column to Drill Down", scale=2)
                    hier_delim = gr.Dropdown(choices=DELIMITER_CHOICES, value=";", label="Separator", scale=1)
                    hier_add_btn = gr.Button("+ Add Level", elem_classes=["btn-action"], scale=1)
 
                with gr.Row():
                    hier_show_class_l1 = gr.Checkbox(label="Show Column Name in Level 1", value=False)
                    hier_show_class_deep = gr.Checkbox(label="Show Column Name in Deeper Levels", value=False)
 
                hier_data = gr.Dataframe(headers=["Level", "Column Name", "Separator"], datatype=["number", "str", "str"], row_count=0, col_count=(3, "fixed"), interactive=False, label="Structure Preview")
                with gr.Row(elem_classes=["clean-row"]):
                    hier_up_btn = gr.Button("↑ Move Up", elem_classes=["btn-action"])
                    hier_down_btn = gr.Button("↓ Move Down", elem_classes=["btn-action"])
                    hier_del_btn = gr.Button("× Delete Selected", elem_classes=["btn-del"])
 
        # --- ADVANCED ---
        with gr.Column(scale=2, visible=False) as advanced_group:
            with gr.Group(elem_classes=["step-card"]):
                gr.HTML('<div class="step-header" style="color:#7c3aed; border-color:#ddd6fe;">Advanced</div>')
                # 1. Select Entities
                gr.Markdown("### 1. Select Entities")
                sheets_sel = gr.CheckboxGroup(choices=[], label="Include Sheets")
                gr.HTML('<div style="height:12px"></div>')
 
                # 2. Node Title
                gr.Markdown("### 2. Node Title Configuration")
                with gr.Row(elem_classes=["clean-row"]):
                    nt_sheet = gr.Dropdown(choices=[], label="Sheet", scale=2)
                    nt_col = gr.Dropdown(choices=[], label="Title Column", scale=2)
                    nt_btn = gr.Button("Set Title", elem_classes=["btn-action"], scale=1)
                nt_view = gr.Dataframe(headers=["Sheet", "Title Column"], datatype=["str", "str"], interactive=False)
                gr.HTML('<div style="height:12px"></div>')
 
                # 3. Direction
                direction = gr.Radio(choices=["left", "right"], value="left", label="Edge Direction")
                gr.HTML('<div style="height:12px"></div>')
 
                # 4. Grouping
                gr.Markdown("### 4. Grouping Rules")
                with gr.Row(elem_classes=["clean-row"]):
                    grp_name = gr.Textbox(label="Name", scale=2)
                    grp_entity = gr.Dropdown(choices=[], label="Entity", scale=2)
                    grp_col = gr.Dropdown(choices=[], label="Column", scale=2)
                    grp_delim = gr.Dropdown(choices=DELIMITER_CHOICES, value=";", label="Separator", scale=1)
                    grp_add_btn = gr.Button("Add", elem_classes=["btn-action"], scale=1)
                grp_data = gr.Dataframe(headers=["Group Name", "Entity", "Column", "Separator"], datatype=["str", "str", "str", "str"], row_count=0, col_count=(4, "fixed"), interactive=False)
                grp_del_btn = gr.Button("Delete Group", elem_classes=["btn-del"])
                gr.HTML('<div style="height:12px"></div>')
 
                # 5. Types
                gr.Markdown("### 5. Column Types")
                types_tab_sheet = gr.Radio(choices=[], label="Sheet")
                types_state = gr.State({})
                types_html = gr.HTML()
 
    # --- STEP 3 ---
    with gr.Group(elem_classes=["step-card"]):
        gr.HTML('<div class="step-header">Step 3: Generate</div>')
        gen_btn = gr.Button("🚀 Generate XML Graph", variant="primary", elem_classes=["btn-main"])
        gen_file = gr.File(label="Download Result", visible=False)
        gen_status = gr.Markdown()
 
    # --- LOGIC ---
    def format_hier_df(sheet_name):
        if not sheet_name or sheet_name not in conv.hierarchy_configs: return []
        res = []
        for i, item in enumerate(conv.hierarchy_configs[sheet_name]):
            if isinstance(item, dict):
                res.append([i+1, item['col'], item.get('delim', ';')])
            else:
                res.append([i+1, str(item), ';'])
        return res
 
    def format_grp_df():
        d = []
        for r in conv.manual_grouping_rules:
            n = r.get("name", "Group")
            delim = r.get("delimiter", ";")
            for e, c in r.get("pairs", []): d.append([n, e, c, delim])
        return d
 
    def format_nt_df():
        return [[s, c] for s, c in conv.sheet_title_cols.items()]
 
    def invalidate(): return gr.update(value=None, visible=False)
 
    def toggle_mode(mode): return gr.update(visible=(mode == "Advanced"))
 
    def update_ui_after_load(ok, msg, sheets):
        if not ok:
            return msg, gr.update(choices=[]), gr.update(choices=[]), gr.update(choices=[]), gr.update(choices=[]), gr.update(choices=[]), "", [], [], invalidate(), gr.update(value=True), gr.update(choices=[]), gr.update(choices=[]), gr.update(choices=[]), []
 
        conv.selected_sheets = set([s for s in sheets if s != "structure"])
        sel_list = sorted(list(conv.selected_sheets))
        first = sel_list[0] if sel_list else None
        new_root_state = (len(sel_list) > 1)
        conv.use_file_root = new_root_state
 
        for s in sel_list:
            if s in conv.sheets_data and not conv.sheet_title_cols.get(s):
                cols = list(conv.sheets_data[s].columns)
                if cols: conv.sheet_title_cols[s] = cols[0]
        t_str = ""
        if first and first in conv.sheets_data:
            t_str = render_types_table(first, {}, list(conv.sheets_data[first].columns))
        return (msg, gr.update(choices=sheets, value=list(conv.selected_sheets)), gr.update(choices=sel_list, value=first), gr.update(choices=sel_list, value=first), gr.update(choices=sel_list, value=first), gr.update(choices=sel_list, value=first), t_str, [], [], invalidate(), gr.update(value=new_root_state), gr.update(choices=sel_list, value=first), gr.update(choices=list(conv.sheets_data[first].columns) if first else []), gr.update(choices=sel_list, value=first), format_nt_df())
 
    def do_load_file(file):
        if not file: return update_ui_after_load(False, "No file", [])
        ok, msg, ch = conv.load_xlsx(file.name)
        return update_ui_after_load(ok, msg, ch)
 
    def do_load_gs(url, json_txt):
        ok, msg, ch = conv.load_gsheet(url, json_txt)
        return update_ui_after_load(ok, msg, ch)
 
    def set_root_flag(val): conv.use_file_root = val; return invalidate()
 
    def set_direction(val): conv.edge_direction = val; return invalidate()
 
    # NEW: handler for exclude-title checkbox
    def set_exclude_title(val):
        conv.exclude_title_from_data = val
        return invalidate()
 
    def on_hier_sheet_change(sheet):
        _auto_save()  # persist before switching to a different sheet
        if not sheet or sheet not in conv.sheets_data: return gr.update(choices=[]), []
        cols = list(conv.sheets_data[sheet].columns)
        current = conv.hierarchy_configs.get(sheet, [])
        used_cols = [x['col'] if isinstance(x, dict) else str(x) for x in current]
        avail = [c for c in cols if c not in used_cols]
        return gr.update(choices=avail), format_hier_df(sheet)
 
    def add_hier_level(sheet, col, delim):
        if not sheet or not col: return format_hier_df(sheet), gr.update(), invalidate()
        current = conv.hierarchy_configs.get(sheet, [])
        used_cols = [x['col'] if isinstance(x, dict) else str(x) for x in current]
        if col in used_cols:
            # Duplicate — silently reject, return unchanged state
            cols = list(conv.sheets_data[sheet].columns)
            avail = [c for c in cols if c not in used_cols]
            return format_hier_df(sheet), gr.update(choices=avail, value=None), invalidate()
        d = None if delim == "No Separation" else delim
        current.append({'col': col, 'delim': d})
        conv.hierarchy_configs[sheet] = current
        cols = list(conv.sheets_data[sheet].columns)
        used_cols = [x['col'] if isinstance(x, dict) else str(x) for x in current]
        avail = [c for c in cols if c not in used_cols]
        _auto_save()
        return format_hier_df(sheet), gr.update(choices=avail, value=None), invalidate()
 
    hier_sel = gr.State(-1)
    grp_sel = gr.State(-1)
 
    def on_hier_select(evt: gr.SelectData): return evt.index[0]
 
    def hier_delete(sheet, idx):
        if not sheet or idx < 0: return format_hier_df(sheet), gr.update(), invalidate()
        current = conv.hierarchy_configs.get(sheet, [])
        if idx < len(current):
            current.pop(idx)
            conv.hierarchy_configs[sheet] = current
        cols = list(conv.sheets_data[sheet].columns)
        used_cols = [x['col'] if isinstance(x, dict) else str(x) for x in current]
        avail = [c for c in cols if c not in used_cols]
        _auto_save()
        return format_hier_df(sheet), gr.update(choices=avail), invalidate()
 
    def hier_move(sheet, idx, direction):
        if not sheet: return format_hier_df(sheet), invalidate()
        cur = conv.hierarchy_configs.get(sheet, [])
        if not cur: return format_hier_df(sheet), invalidate()
        if direction == "up" and idx > 0: cur[idx], cur[idx-1] = cur[idx-1], cur[idx]
        elif direction == "down" and idx < len(cur)-1: cur[idx], cur[idx+1] = cur[idx+1], cur[idx]
        conv.hierarchy_configs[sheet] = cur
        _auto_save()
        return format_hier_df(sheet), invalidate()
 
    def set_l1(v): conv.use_prefix_l1 = v; return invalidate()
 
    def set_deep(v): conv.use_prefix_deep = v; return invalidate()
 
    def add_grp(name, entity, col, delim):
        if not entity or not col: return format_grp_df(), invalidate()
        d = None if delim == "No Separation" else delim
        conv.manual_grouping_rules.append({"name": name or "Group", "pairs": [(entity, col)], "delimiter": d})
        return format_grp_df(), invalidate()
 
    def on_grp_select(evt: gr.SelectData): return evt.index[0]
 
    def del_grp(idx):
        if 0 <= idx < len(conv.manual_grouping_rules): conv.manual_grouping_rules.pop(idx)
        return format_grp_df(), invalidate()
 
    def on_save_sel(sel):
        conv.selected_sheets = set(sel or [])
        new_choices = sorted(list(conv.selected_sheets))
        conv.hierarchy_configs = {k: v for k, v in conv.hierarchy_configs.items() if k in new_choices}
        conv.sheet_title_cols = {k: v for k, v in conv.sheet_title_cols.items() if k in new_choices}
        return (gr.update(choices=new_choices, value=None), gr.update(choices=new_choices, value=None), gr.update(choices=new_choices, value=None), "", invalidate(), gr.update(choices=new_choices, value=None), format_nt_df())
 
    def on_types_tab(sheet, t_state):
        if not sheet or sheet not in conv.sheets_data: return render_types_table(sheet, t_state, [])
        return render_types_table(sheet, t_state, list(conv.sheets_data[sheet].columns))
 
    def render_types_table(sheet, state, cols):
        if not sheet: return "Select sheet"
        rows = []
        prefs = state.get(sheet, {})
        for c in cols:
            cur = prefs.get(c, "auto")
            opt = "".join(f'<option {"selected" if cur==t else ""}>{t}</option>' for t in TYPE_CHOICES)
            rows.append(f"<tr><td>{c}</td><td><select>{opt}</select></td></tr>")
        return f"<table>{chr(10).join(rows)}</table>"
 
    def on_nt_sheet_change(sheet):
        if not sheet or sheet not in conv.sheets_data: return gr.update(choices=[])
        return gr.update(choices=list(conv.sheets_data[sheet].columns))
 
    def set_node_title(sheet, col):
        if sheet and col: conv.sheet_title_cols[sheet] = col
        return format_nt_df(), invalidate()
 
    def gen_xml(mode_val, use_l1, use_deep, data_delim):
        conv.use_prefix_l1 = use_l1
        conv.use_prefix_deep = use_deep
        conv.data_delimiter = None if data_delim == "No Separation" else data_delim
        try:
            conv.apply_groupings()
            path, n, m = conv.generate(mode_val)
            return gr.update(value=path, visible=True), f"✅ Success! {n} nodes, {m} edges generated."
        except Exception as e:
            traceback.print_exc()
            return gr.update(visible=False), f"❌ **Error:** {str(e)}"

    # ── TEMPLATE HANDLERS ──────────────────────────────────────

    def tpl_init_save(login):
        if not (login or "").strip():
            return gr.update(visible=False), gr.update(visible=False), gr.update(visible=False, value=False), "⚠️ Please enter a login before saving."
        if not conv.loaded_sheets:
            return gr.update(visible=False), gr.update(visible=False), gr.update(visible=False, value=False), "⚠️ Please load a file before saving a template."
        return gr.update(visible=True), gr.update(visible=False), gr.update(visible=False, value=False), ""

    def tpl_confirm_save(login, name, overwrite):
        login = (login or "").strip()
        if not login:
            return gr.update(visible=True), gr.update(visible=False), gr.update(), "⚠️ Login is required."
        if not conv.loaded_sheets:
            return gr.update(visible=False), gr.update(visible=False), gr.update(), "⚠️ No file loaded — cannot save template."
        name = (name or "").strip()
        if not name:
            name = f"template_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        all_tpls = _tmpl_load_all()
        user_tpls = all_tpls.get(login, {})
        if name in user_tpls and not overwrite:
            return (gr.update(visible=True), gr.update(visible=True), gr.update(),
                    f"⚠️ Template **{name}** already exists. Check 'Overwrite if exists' and click Confirm again.")
        action_word = "overwritten" if (name in user_tpls and overwrite) else "saved"
        entry = {
            "created": datetime.datetime.now().isoformat(),
            "structural_metadata": _tmpl_snapshot_structure(),
            "config": _tmpl_collect_config(),
        }
        user_tpls[name] = entry
        all_tpls[login] = user_tpls
        _tmpl_save_all(all_tpls)
        return gr.update(visible=False), gr.update(visible=False), gr.update(value=""), f"✅ Template **{name}** {action_word} successfully."

    def tpl_init_load(login):
        login = (login or "").strip()
        if not login:
            return gr.update(visible=False), gr.update(visible=False), gr.update(choices=[]), "⚠️ Please enter a login before loading."
        if not conv.loaded_sheets:
            return gr.update(visible=False), gr.update(visible=False), gr.update(choices=[]), "⚠️ Please load a file before applying a template."
        all_tpls = _tmpl_load_all()
        user_tpls = all_tpls.get(login, {})
        if not user_tpls:
            return gr.update(visible=False), gr.update(visible=False), gr.update(choices=[]), f"ℹ️ No templates found for login **{login}**."
        names = sorted(user_tpls.keys())
        return gr.update(visible=True), gr.update(visible=False), gr.update(choices=names, value=names[0]), ""

    def tpl_apply_fn(login, tpl_name):
        # 15 outputs: tpl_status, tpl_load_panel, sheets_sel, hier_sheet, hier_data,
        # hier_add_col, nt_view, grp_data, direction, use_file_root_cb,
        # exclude_title_cb, global_data_delim, gen_file, tpl_summary_accordion, tpl_summary

        def _noop(msg):
            return (msg, gr.update(), gr.update(), gr.update(), gr.update(), gr.update(),
                    gr.update(), gr.update(), gr.update(), gr.update(), gr.update(),
                    gr.update(), gr.update(), gr.update(visible=False), gr.update())

        login = (login or "").strip()
        if not login or not tpl_name:
            return _noop("⚠️ Please select a template to apply.")
        all_tpls = _tmpl_load_all()
        template = all_tpls.get(login, {}).get(tpl_name)
        if not template:
            return _noop(f"⚠️ Template **{tpl_name}** not found for login **{login}**.")

        applied, skipped, match_type = _tmpl_apply(template)

        if match_type == "none":
            return _noop(
                f"❌ Template **{tpl_name}** is incompatible with the current file — "
                f"fewer than half of the structural settings could be applied."
            )

        # Build summary markdown
        parts = []
        if applied:
            parts.append("**Applied:**\n" + "\n".join(f"- {a}" for a in applied))
        if skipped:
            parts.append("**Skipped:**\n" + "\n".join(f"- {s}" for s in skipped))
        summary_md = "\n\n".join(parts)

        if match_type == "full":
            status_msg = f"✅ Template **{tpl_name}** applied — {len(applied)} settings applied."
        else:
            status_msg = (f"⚠️ Template **{tpl_name}** partially applied — "
                          f"{len(applied)} applied, {len(skipped)} skipped.")

        sel_list = sorted(list(conv.selected_sheets))
        first = sel_list[0] if sel_list else None
        hier_avail = []
        if first and first in conv.sheets_data:
            all_c = list(conv.sheets_data[first].columns)
            used_c = [lv["col"] if isinstance(lv, dict) else str(lv) for lv in conv.hierarchy_configs.get(first, [])]
            hier_avail = [c for c in all_c if c not in used_c]

        delim_val = conv.data_delimiter if conv.data_delimiter else "No Separation"

        return (
            status_msg,                                        # tpl_status
            gr.update(visible=False),                          # tpl_load_panel
            gr.update(value=sel_list),                         # sheets_sel
            gr.update(value=first),                            # hier_sheet
            format_hier_df(first),                             # hier_data
            gr.update(choices=hier_avail),                     # hier_add_col
            format_nt_df(),                                    # nt_view
            format_grp_df(),                                   # grp_data
            gr.update(value=conv.edge_direction),              # direction
            gr.update(value=conv.use_file_root),               # use_file_root_cb
            gr.update(value=conv.exclude_title_from_data),     # exclude_title_cb
            gr.update(value=delim_val),                        # global_data_delim
            invalidate(),                                      # gen_file
            gr.update(visible=bool(summary_md)),               # tpl_summary_accordion
            summary_md,                                        # tpl_summary
        )

    # --- BINDINGS ---
    mode_toggle.change(toggle_mode, inputs=[mode_toggle], outputs=[advanced_group])
    use_file_root_cb.change(set_root_flag, inputs=[use_file_root_cb], outputs=[gen_file])
    # NEW binding
    exclude_title_cb.change(set_exclude_title, inputs=[exclude_title_cb], outputs=[gen_file])
    direction.change(set_direction, inputs=[direction], outputs=[gen_file])
    common_outputs = [status, sheets_sel, hier_sheet, grp_entity, types_tab_sheet, grp_entity, types_html, hier_data, grp_data, gen_file, use_file_root_cb, nt_sheet, nt_col, types_tab_sheet, nt_view]
    load_file_btn.click(do_load_file, inputs=[up_file], outputs=common_outputs)
    load_gs_btn.click(do_load_gs, inputs=[gs_url, gs_json], outputs=common_outputs)
 
    sheets_sel.change(on_save_sel, inputs=[sheets_sel], outputs=[hier_sheet, grp_entity, types_tab_sheet, types_html, gen_file, nt_sheet, nt_view])
    hier_sheet.change(on_hier_sheet_change, inputs=[hier_sheet], outputs=[hier_add_col, hier_data])
    hier_add_btn.click(add_hier_level, inputs=[hier_sheet, hier_add_col, hier_delim], outputs=[hier_data, hier_add_col, gen_file])
 
    hier_data.select(on_hier_select, None, hier_sel)
    hier_up_btn.click(lambda s, i: hier_move(s, i, "up"), inputs=[hier_sheet, hier_sel], outputs=[hier_data, gen_file])
    hier_down_btn.click(lambda s, i: hier_move(s, i, "down"), inputs=[hier_sheet, hier_sel], outputs=[hier_data, gen_file])
    hier_del_btn.click(hier_delete, inputs=[hier_sheet, hier_sel], outputs=[hier_data, hier_add_col, gen_file])
 
    hier_show_class_l1.change(set_l1, inputs=[hier_show_class_l1], outputs=[gen_file])
    hier_show_class_deep.change(set_deep, inputs=[hier_show_class_deep], outputs=[gen_file])
    global_data_delim.change(invalidate, outputs=[gen_file])
 
    grp_entity.change(lambda s: gr.update(choices=list(conv.sheets_data[s].columns) if s in conv.sheets_data else []), inputs=[grp_entity], outputs=[grp_col])
    grp_add_btn.click(add_grp, inputs=[grp_name, grp_entity, grp_col, grp_delim], outputs=[grp_data, gen_file])
    grp_data.select(on_grp_select, None, grp_sel)
    grp_del_btn.click(del_grp, inputs=[grp_sel], outputs=[grp_data, gen_file])
 
    nt_sheet.change(on_nt_sheet_change, inputs=[nt_sheet], outputs=[nt_col])
    nt_btn.click(set_node_title, inputs=[nt_sheet, nt_col], outputs=[nt_view, gen_file])
    types_tab_sheet.change(on_types_tab, inputs=[types_tab_sheet, types_state], outputs=[types_html])
    gen_btn.click(gen_xml, inputs=[mode_toggle, hier_show_class_l1, hier_show_class_deep, global_data_delim], outputs=[gen_file, gen_status])

    # ── TEMPLATE BINDINGS ──────────────────────────────────────
    tpl_save_btn.click(
        tpl_init_save,
        inputs=[tpl_login],
        outputs=[tpl_save_panel, tpl_load_panel, tpl_overwrite_cb, tpl_status],
    )
    tpl_confirm_save_btn.click(
        tpl_confirm_save,
        inputs=[tpl_login, tpl_name_input, tpl_overwrite_cb],
        outputs=[tpl_save_panel, tpl_overwrite_cb, tpl_name_input, tpl_status],
    )
    tpl_load_btn.click(
        tpl_init_load,
        inputs=[tpl_login],
        outputs=[tpl_load_panel, tpl_save_panel, tpl_list, tpl_status],
    )
    tpl_confirm_load_btn.click(
        tpl_apply_fn,
        inputs=[tpl_login, tpl_list],
        outputs=[
            tpl_status, tpl_load_panel,
            sheets_sel, hier_sheet, hier_data, hier_add_col,
            nt_view, grp_data, direction, use_file_root_cb,
            exclude_title_cb, global_data_delim, gen_file,
            tpl_summary_accordion, tpl_summary,
        ],
    )

    app.load(None, None, gs_json, js="() => localStorage.getItem('gs_creds') || ''")
    gs_json.change(None, gs_json, None, js="(v) => localStorage.setItem('gs_creds', v)")
 
if __name__ == "__main__":
    app.launch()