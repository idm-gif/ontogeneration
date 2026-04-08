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
.adv-table { width:100%; border-collapse:collapse; font-size:13px; margin-top:8px; }
.adv-table th { text-align:left; padding:8px 10px; background:#f3f4f6; border:1px solid #e5e7eb; font-weight:600; color:#374151; white-space:nowrap; }
.adv-table td { padding:6px 10px; border:1px solid #e5e7eb; vertical-align:middle; }
.adv-table tr:nth-child(even) td { background:#fafafa; }
.adv-table tr:hover td { background:#eff6ff; }
.adv-col-name { font-weight:600; color:#1f2937; min-width:120px; }
.adv-table select { width:100%; border:1px solid #d1d5db; border-radius:4px; padding:4px 6px; font-size:12px; background:#fff; cursor:pointer; }
.adv-table select:focus { outline:2px solid #2563eb; border-color:#2563eb; }
.group-tag { display:inline-flex; align-items:center; gap:6px; background:#ede9fe; color:#5b21b6; padding:4px 12px; border-radius:20px; margin:3px; font-size:13px; font-weight:500; }
.group-tag button { background:none; border:none; color:#7c3aed; cursor:pointer; font-size:15px; line-height:1; padding:0; }
.sheet-btn-row { display:flex; flex-wrap:wrap; gap:8px; margin-bottom:8px; }
.sheet-btn { padding:6px 16px; border-radius:8px; border:2px solid #e5e7eb; background:#fff; cursor:pointer; font-size:13px; font-weight:500; color:#374151; transition:all 0.15s; }
.sheet-btn.active { border-color:#2563eb; background:#eff6ff; color:#2563eb; }
/* Sheet selector — browser-style tabs */
#adv-sheet-radio { margin-bottom:0 !important; }
#adv-sheet-radio > div > div { border-bottom:2px solid #e5e7eb !important; display:flex !important; flex-wrap:nowrap !important; overflow-x:auto !important; gap:0 !important; padding-bottom:0 !important; align-items:flex-end !important; }
#adv-sheet-radio label { display:inline-flex !important; align-items:center !important; justify-content:center !important; padding:8px 20px !important; border:1px solid #e5e7eb !important; border-bottom:none !important; border-radius:8px 8px 0 0 !important; margin-right:3px !important; margin-bottom:-2px !important; cursor:pointer !important; background:#f3f4f6 !important; color:#6b7280 !important; font-size:13px !important; font-weight:500 !important; white-space:nowrap !important; transition:background 0.15s, color 0.15s !important; position:relative !important; z-index:1 !important; }
#adv-sheet-radio label:hover { background:#e9ecef !important; color:#374151 !important; }
#adv-sheet-radio label:has(input:checked) { background:#fff !important; color:#2563eb !important; border-color:#c7d2fe #c7d2fe #fff !important; font-weight:600 !important; box-shadow:0 -2px 0 0 #2563eb inset !important; }
#adv-sheet-radio input[type=radio] { display:none !important; }
#adv-sheet-radio > div > span { display:none !important; }
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
 
    # --- STEP 2 (Simple Mode) ---
    with gr.Row(visible=True) as step2_simple_row:
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

    # --- ADVANCED MODE (Full-Width) ---
    with gr.Group(elem_classes=["step-card"], visible=False) as advanced_group:
        gr.HTML('<div class="step-header" style="color:#7c3aed; border-color:#ddd6fe;">Advanced Configuration</div>')

        # State for advanced mode config
        adv_state = gr.State({"groups": [], "sheets": {}})
        adv_cur_sheet = gr.State("")
        # Hidden textbox: JS writes column-config changes here
        adv_col_change = gr.Textbox(visible=False, elem_id="adv-col-change-tb")

        # --- Included Sheets ---
        gr.Markdown("**Included Sheets**", elem_classes=["clean-row"])
        sheets_sel = gr.CheckboxGroup(choices=[], label="", elem_id="adv-sheets-sel")

        # --- Edge Direction ---
        direction = gr.Radio(choices=["left", "right"], value="left", label="Edge Direction")

        gr.HTML('<hr style="border:none;border-top:1px solid #e5e7eb;margin:12px 0">')

        # --- Cross-Sheet Groups ---
        gr.Markdown("### Cross-Sheet Groups")
        gr.Markdown("Create named groups — then assign them to columns in any sheet's table below.")
        with gr.Row(elem_classes=["clean-row"]):
            adv_grp_input = gr.Textbox(label="New Group Name", placeholder="e.g. Fruits", scale=3)
            adv_grp_add_btn = gr.Button("+ Add Group", elem_classes=["btn-action"], scale=1)
            adv_grp_del_dropdown = gr.Dropdown(choices=[], label="Delete Group", scale=2)
            adv_grp_del_btn = gr.Button("× Delete", elem_classes=["btn-del"], scale=1)
        adv_grp_display = gr.HTML("<p style='color:#888;font-size:13px'>No groups yet.</p>")

        gr.HTML('<hr style="border:none;border-top:1px solid #e5e7eb;margin:12px 0">')

        # --- Per-Sheet Configuration ---
        gr.Markdown("### Sheet Configuration")
        gr.Markdown("Select a sheet to configure its columns — settings are saved per sheet.")

        adv_sheet_radio = gr.Radio(choices=[], label="", elem_id="adv-sheet-radio")

        with gr.Row(elem_classes=["clean-row"]):
            adv_num_levels = gr.Dropdown(
                choices=[0, 1, 2, 3, 4, 5], value=0,
                label="Number of Hierarchy Levels",
                info="How many hierarchy levels are selectable in the table below.",
                scale=2
            )
            adv_title_col = gr.Dropdown(choices=[], label="Title Column (node name)", scale=3)

        # The interactive configuration table (rendered as HTML with JS)
        adv_table_html = gr.HTML("<p style='color:#888'>Load a file and select a sheet.</p>")

        # Legacy components kept for simple-mode compatibility (hidden)
        nt_sheet = gr.Dropdown(choices=[], visible=False)
        nt_col = gr.Dropdown(choices=[], visible=False)
        nt_btn = gr.Button(visible=False)
        nt_view = gr.Dataframe(headers=["Sheet", "Title Column"], datatype=["str", "str"], interactive=False, visible=False)
        grp_name = gr.Textbox(visible=False)
        grp_entity = gr.Dropdown(choices=[], visible=False)
        grp_col = gr.Dropdown(choices=[], visible=False)
        grp_delim = gr.Dropdown(choices=DELIMITER_CHOICES, value=";", visible=False)
        grp_add_btn = gr.Button(visible=False)
        grp_data = gr.Dataframe(headers=["Group Name", "Entity", "Column", "Separator"], datatype=["str", "str", "str", "str"], row_count=0, col_count=(4, "fixed"), interactive=False, visible=False)
        grp_del_btn = gr.Button(visible=False)
        types_tab_sheet = gr.Radio(choices=[], visible=False)
        types_state = gr.State({})
        types_html = gr.HTML(visible=False)
 
    # --- STEP 3 ---
    with gr.Group(elem_classes=["step-card"]):
        gr.HTML('<div class="step-header">Step 3: Generate</div>')
        gen_btn = gr.Button("🚀 Generate XML Graph", variant="primary", elem_classes=["btn-main"])
        gen_file = gr.File(label="Download Result", visible=False)
        gen_status = gr.Markdown()
 
    # --- LOGIC ---

    # ── Simple-mode helpers ───────────────────────────────────────
    def format_hier_df(sheet_name):
        if not sheet_name or sheet_name not in conv.hierarchy_configs: return []
        res = []
        for i, item in enumerate(conv.hierarchy_configs[sheet_name]):
            if isinstance(item, dict):
                res.append([i+1, item['col'], item.get('delim', ';')])
            else:
                res.append([i+1, str(item), ';'])
        return res

    def invalidate(): return gr.update(value=None, visible=False)

    # ── Advanced-mode helpers ─────────────────────────────────────

    def render_adv_groups(adv_config):
        groups = adv_config.get("groups", [])
        if not groups:
            return "<p style='color:#888;font-size:13px;margin:4px 0'>No groups yet. Add a name above.</p>"
        tags = "".join(
            f'<span class="group-tag">{g}</span>' for g in groups
        )
        return f'<div style="display:flex;flex-wrap:wrap;gap:4px;margin-top:4px">{tags}</div>'

    def render_adv_table(sheet, adv_config):
        if not sheet or sheet not in conv.sheets_data:
            return "<p style='color:#888'>Load a file and select a sheet to configure.</p>"

        sheet_cfg = adv_config.get("sheets", {}).get(sheet, {})
        num_levels = int(sheet_cfg.get("num_levels", 0))
        col_configs = sheet_cfg.get("col_configs", {})
        groups = adv_config.get("groups", [])
        cols = list(conv.sheets_data[sheet].columns)

        hier_header = f"Hierarchy (1–{num_levels})" if num_levels > 0 else "Hierarchy Level"

        rows_html = []
        for col in cols:
            cfg = col_configs.get(col, {})
            hier_val = cfg.get("hierarchy")
            sep_val  = cfg.get("separator", ";")
            type_val = cfg.get("type", "auto")
            grp_val  = cfg.get("group") or ""

            # Hierarchy cell
            if num_levels > 0:
                h_opts = '<option value="">None</option>'
                for i in range(1, num_levels + 1):
                    sel = "selected" if hier_val == i else ""
                    h_opts += f'<option value="{i}" {sel}>Level {i}</option>'
                hier_cell = f'<select class="hier-sel" onchange="advHierChanged(this)">{h_opts}</select>'
            else:
                hier_cell = '<span style="color:#bbb;font-size:12px">set levels above</span>'

            # Separator cell
            sep_cell = '<select class="sep-sel" onchange="advTableChanged()">'
            for s in [";", ",", "|", "No Separation"]:
                sel = "selected" if sep_val == s else ""
                sep_cell += f'<option value="{s}" {sel}>{s}</option>'
            sep_cell += "</select>"

            # Type cell
            type_cell = '<select class="type-sel" onchange="advTableChanged()">'
            for t in TYPE_CHOICES:
                sel = "selected" if type_val == t else ""
                type_cell += f'<option value="{t}" {sel}>{t}</option>'
            type_cell += "</select>"

            # Group cell
            grp_cell = '<select class="grp-sel" onchange="advTableChanged()">'
            grp_cell += '<option value="">None</option>'
            for g in groups:
                sel = "selected" if grp_val == g else ""
                grp_cell += f'<option value="{g}" {sel}>{g}</option>'
            grp_cell += "</select>"

            safe_col = col.replace('"', '&quot;')
            rows_html.append(
                f'<tr data-col="{safe_col}">'
                f'<td class="adv-col-name">{col}</td>'
                f'<td>{hier_cell}</td>'
                f'<td>{sep_cell}</td>'
                f'<td>{type_cell}</td>'
                f'<td>{grp_cell}</td>'
                f'</tr>'
            )

        js_block = f'''<script>
(function(){{
  var SHEET = {json.dumps(sheet)};
  function advHierChanged(sel){{
    var val = sel.value;
    if(val){{
      document.querySelectorAll("#adv-config-table .hier-sel").forEach(function(s){{
        if(s !== sel && s.value === val) s.value = "";
      }});
    }}
    advTableChanged();
  }}
  function advTableChanged(){{
    var rows = document.querySelectorAll("#adv-config-table tbody tr[data-col]");
    var colCfgs = {{}};
    rows.forEach(function(row){{
      var col     = row.dataset.col;
      var hierSel = row.querySelector(".hier-sel");
      var sepSel  = row.querySelector(".sep-sel");
      var typeSel = row.querySelector(".type-sel");
      var grpSel  = row.querySelector(".grp-sel");
      colCfgs[col] = {{
        hierarchy: (hierSel && hierSel.value) ? parseInt(hierSel.value) : null,
        separator: sepSel  ? sepSel.value  : ";",
        type:      typeSel ? typeSel.value : "auto",
        group:     (grpSel && grpSel.value) ? grpSel.value : null
      }};
    }});
    var payload = JSON.stringify({{sheet: SHEET, col_configs: colCfgs}});
    var tb = document.querySelector("#adv-col-change-tb textarea, #adv-col-change-tb input");
    if(tb){{
      tb.value = payload;
      tb.dispatchEvent(new Event("input",  {{bubbles:true}}));
      tb.dispatchEvent(new Event("change", {{bubbles:true}}));
    }}
  }}
  window.advHierChanged  = advHierChanged;
  window.advTableChanged = advTableChanged;
}})();
</script>'''

        return (
            js_block
            + f'<table id="adv-config-table" class="adv-table">'
            + f'<thead><tr>'
            + f'<th>Column</th><th>{hier_header}</th>'
            + f'<th>Separator</th><th>Type</th><th>Group</th>'
            + f'</tr></thead>'
            + f'<tbody>{"".join(rows_html)}</tbody>'
            + f'</table>'
        )

    def _init_adv_config(sel_list, prev_config=None):
        import copy
        adv_config = copy.deepcopy(prev_config) if prev_config else {"groups": [], "sheets": {}}
        adv_config.setdefault("groups", [])
        adv_config.setdefault("sheets", {})
        for sheet in sel_list:
            if sheet not in adv_config["sheets"] and sheet in conv.sheets_data:
                cols = list(conv.sheets_data[sheet].columns)
                adv_config["sheets"][sheet] = {
                    "num_levels": 0,
                    "title_col": cols[0] if cols else "",
                    "col_configs": {
                        col: {"hierarchy": None, "separator": ";", "type": "auto", "group": None}
                        for col in cols
                    }
                }
        return adv_config

    def _sync_adv_to_conv(adv_config):
        # 1. hierarchy_configs
        new_hier = {}
        for sheet, scfg in adv_config.get("sheets", {}).items():
            hier_cols = [
                (col, cfg["hierarchy"], cfg.get("separator", ";"))
                for col, cfg in scfg.get("col_configs", {}).items()
                if cfg.get("hierarchy") is not None
            ]
            hier_cols.sort(key=lambda x: x[1])
            if hier_cols:
                new_hier[sheet] = [{"col": c, "delim": s} for c, _, s in hier_cols]
        conv.hierarchy_configs = new_hier

        # 2. column_type_preferences
        new_ctp = {}
        for sheet, scfg in adv_config.get("sheets", {}).items():
            prefs = {
                col: cfg["type"]
                for col, cfg in scfg.get("col_configs", {}).items()
                if cfg.get("type", "auto") != "auto"
            }
            if prefs:
                new_ctp[sheet] = prefs
        conv.column_type_preferences = new_ctp

        # 3. sheet_title_cols
        new_stc = {}
        for sheet, scfg in adv_config.get("sheets", {}).items():
            tc = scfg.get("title_col", "")
            if tc and sheet in conv.sheets_data and tc in conv.sheets_data[sheet].columns:
                new_stc[sheet] = tc
            elif sheet in conv.sheets_data:
                cols = list(conv.sheets_data[sheet].columns)
                if cols:
                    new_stc[sheet] = cols[0]
        conv.sheet_title_cols = new_stc

        # 4. manual_grouping_rules from group column assignments
        group_pairs = {}
        for sheet, scfg in adv_config.get("sheets", {}).items():
            for col, cfg in scfg.get("col_configs", {}).items():
                grp = cfg.get("group")
                if grp:
                    group_pairs.setdefault(grp, []).append(
                        (sheet, col, cfg.get("separator", ";"))
                    )
        conv.manual_grouping_rules = [
            {"name": grp, "pairs": [(s, c) for s, c, _ in pairs],
             "delimiter": pairs[0][2] if pairs else ";"}
            for grp, pairs in group_pairs.items()
        ]

    def _rebuild_adv_from_conv(sel_list):
        """Reconstruct adv_config from conv.* (used after template load)."""
        adv_config = {"groups": [], "sheets": {}}
        # groups from grouping rules
        adv_config["groups"] = list({r.get("name", "Group") for r in conv.manual_grouping_rules})
        # (sheet, col) -> group name
        pair_to_group = {}
        for rule in conv.manual_grouping_rules:
            for s, c in rule.get("pairs", []):
                pair_to_group[(s, c)] = rule.get("name", "Group")

        for sheet in sel_list:
            if sheet not in conv.sheets_data:
                continue
            cols = list(conv.sheets_data[sheet].columns)
            hier_levels = conv.hierarchy_configs.get(sheet, [])
            col_to_level = {}
            col_to_sep   = {}
            for i, lv in enumerate(hier_levels):
                col  = lv["col"] if isinstance(lv, dict) else str(lv)
                sep  = lv.get("delim", ";") if isinstance(lv, dict) else ";"
                col_to_level[col] = i + 1
                col_to_sep[col]   = sep
            type_prefs = conv.column_type_preferences.get(sheet, {})
            title_col  = conv.sheet_title_cols.get(sheet, cols[0] if cols else "")
            col_configs = {}
            for col in cols:
                col_configs[col] = {
                    "hierarchy": col_to_level.get(col),
                    "separator": col_to_sep.get(col, ";"),
                    "type":      type_prefs.get(col, "auto"),
                    "group":     pair_to_group.get((sheet, col)),
                }
            adv_config["sheets"][sheet] = {
                "num_levels": len(hier_levels),
                "title_col":  title_col,
                "col_configs": col_configs,
            }
        return adv_config
 
    # ── Mode toggle ───────────────────────────────────────────────
    def toggle_mode(mode):
        is_adv = (mode == "Advanced")
        return gr.update(visible=not is_adv), gr.update(visible=is_adv)

    # ── Simple-mode event handlers ────────────────────────────────
    def set_root_flag(val): conv.use_file_root = val; return invalidate()
    def set_direction(val): conv.edge_direction = val; return invalidate()
    def set_exclude_title(val): conv.exclude_title_from_data = val; return invalidate()
    def set_l1(v): conv.use_prefix_l1 = v; return invalidate()
    def set_deep(v): conv.use_prefix_deep = v; return invalidate()

    def on_hier_sheet_change(sheet):
        if not sheet or sheet not in conv.sheets_data: return gr.update(choices=[]), []
        cols = list(conv.sheets_data[sheet].columns)
        current = conv.hierarchy_configs.get(sheet, [])
        used = [x['col'] if isinstance(x, dict) else str(x) for x in current]
        return gr.update(choices=[c for c in cols if c not in used]), format_hier_df(sheet)

    def add_hier_level(sheet, col, delim):
        if not sheet or not col: return format_hier_df(sheet), gr.update(), invalidate()
        current = conv.hierarchy_configs.get(sheet, [])
        current.append({'col': col, 'delim': None if delim == "No Separation" else delim})
        conv.hierarchy_configs[sheet] = current
        cols = list(conv.sheets_data[sheet].columns)
        used = [x['col'] if isinstance(x, dict) else str(x) for x in current]
        return format_hier_df(sheet), gr.update(choices=[c for c in cols if c not in used], value=None), invalidate()

    hier_sel = gr.State(-1)

    def on_hier_select(evt: gr.SelectData): return evt.index[0]

    def hier_delete(sheet, idx):
        if not sheet or idx < 0: return format_hier_df(sheet), gr.update(), invalidate()
        current = conv.hierarchy_configs.get(sheet, [])
        if idx < len(current): current.pop(idx)
        conv.hierarchy_configs[sheet] = current
        cols = list(conv.sheets_data[sheet].columns)
        used = [x['col'] if isinstance(x, dict) else str(x) for x in current]
        return format_hier_df(sheet), gr.update(choices=[c for c in cols if c not in used]), invalidate()

    def hier_move(sheet, idx, d):
        if not sheet: return format_hier_df(sheet), gr.update(), invalidate()
        cur = conv.hierarchy_configs.get(sheet, [])
        if not cur: return format_hier_df(sheet), gr.update(), invalidate()
        if d == "up" and idx > 0: cur[idx], cur[idx-1] = cur[idx-1], cur[idx]
        elif d == "down" and idx < len(cur)-1: cur[idx], cur[idx+1] = cur[idx+1], cur[idx]
        conv.hierarchy_configs[sheet] = cur
        return format_hier_df(sheet), gr.update(), invalidate()

    # ── Advanced-mode event handlers ──────────────────────────────

    def on_adv_col_change(change_json, adv_config):
        """JS sends {sheet, col_configs} when any table dropdown changes."""
        if not change_json or not change_json.strip():
            return adv_config
        try:
            payload = json.loads(change_json)
        except Exception:
            return adv_config
        sheet = payload.get("sheet", "")
        col_configs = payload.get("col_configs", {})
        if not sheet:
            return adv_config
        adv_config.setdefault("sheets", {})
        adv_config["sheets"].setdefault(sheet, {"num_levels": 0, "title_col": "", "col_configs": {}})
        # Enforce unique hierarchy levels — last assignment wins
        seen_levels = {}
        for col, cfg in col_configs.items():
            lvl = cfg.get("hierarchy")
            if lvl is not None:
                if lvl in seen_levels:
                    col_configs[seen_levels[lvl]]["hierarchy"] = None
                seen_levels[lvl] = col
        # Preserve num_levels and title_col; only update col_configs
        adv_config["sheets"][sheet]["col_configs"] = col_configs
        return adv_config

    def on_adv_sheet_change(sheet, adv_config):
        """User clicks a different sheet tab — reload controls for that sheet."""
        if not sheet or sheet not in conv.sheets_data:
            return sheet, gr.update(value=0), gr.update(choices=[], value=None), \
                   "<p style='color:#888'>Select a sheet.</p>"
        scfg   = adv_config.get("sheets", {}).get(sheet, {})
        n_lvl  = int(scfg.get("num_levels", 0))
        t_col  = scfg.get("title_col", "")
        cols   = list(conv.sheets_data[sheet].columns)
        t_val  = t_col if t_col in cols else (cols[0] if cols else None)
        return sheet, gr.update(value=n_lvl), gr.update(choices=cols, value=t_val), \
               render_adv_table(sheet, adv_config)

    def on_adv_num_levels(n, sheet, adv_config):
        """User changes hierarchy-level count for the current sheet."""
        if not sheet:
            return adv_config, render_adv_table(sheet, adv_config)
        n = int(n) if n is not None else 0
        adv_config.setdefault("sheets", {})
        adv_config["sheets"].setdefault(sheet, {"num_levels": 0, "title_col": "", "col_configs": {}})
        adv_config["sheets"][sheet]["num_levels"] = n
        # Clear hierarchy assignments above the new max
        for cfg in adv_config["sheets"][sheet].get("col_configs", {}).values():
            if cfg.get("hierarchy") and cfg["hierarchy"] > n:
                cfg["hierarchy"] = None
        return adv_config, render_adv_table(sheet, adv_config)

    def on_adv_title_col(title_col, sheet, adv_config):
        if not sheet or not title_col:
            return adv_config
        adv_config.setdefault("sheets", {})
        adv_config["sheets"].setdefault(sheet, {"num_levels": 0, "title_col": "", "col_configs": {}})
        adv_config["sheets"][sheet]["title_col"] = title_col
        return adv_config

    def on_adv_add_group(grp_name, sheet, adv_config):
        grp_name = (grp_name or "").strip()
        if not grp_name:
            return adv_config, gr.update(), render_adv_groups(adv_config), \
                   gr.update(), render_adv_table(sheet, adv_config)
        adv_config.setdefault("groups", [])
        if grp_name not in adv_config["groups"]:
            adv_config["groups"].append(grp_name)
        groups = adv_config["groups"]
        return (adv_config, gr.update(value=""), render_adv_groups(adv_config),
                gr.update(choices=groups, value=None), render_adv_table(sheet, adv_config))

    def on_adv_del_group(grp_name, sheet, adv_config):
        if not grp_name:
            return adv_config, render_adv_groups(adv_config), gr.update(), \
                   render_adv_table(sheet, adv_config)
        if "groups" in adv_config and grp_name in adv_config["groups"]:
            adv_config["groups"].remove(grp_name)
        # Clear this group from all sheets
        for scfg in adv_config.get("sheets", {}).values():
            for cfg in scfg.get("col_configs", {}).values():
                if cfg.get("group") == grp_name:
                    cfg["group"] = None
        groups = adv_config.get("groups", [])
        return (adv_config, render_adv_groups(adv_config),
                gr.update(choices=groups, value=None), render_adv_table(sheet, adv_config))

    def on_adv_sheets_sel(sel, adv_config):
        """Included-sheets CheckboxGroup changed in Advanced mode."""
        conv.selected_sheets = set(sel or [])
        new_choices = sorted(list(conv.selected_sheets))
        conv.hierarchy_configs   = {k: v for k, v in conv.hierarchy_configs.items()   if k in new_choices}
        conv.sheet_title_cols    = {k: v for k, v in conv.sheet_title_cols.items()    if k in new_choices}
        adv_config = _init_adv_config(new_choices, adv_config)
        first = new_choices[0] if new_choices else None
        cols  = list(conv.sheets_data[first].columns) if first and first in conv.sheets_data else []
        t_val = adv_config.get("sheets", {}).get(first, {}).get("title_col", cols[0] if cols else None)
        return (
            gr.update(choices=new_choices, value=first),          # hier_sheet (simple)
            adv_config,                                            # adv_state
            first or "",                                           # adv_cur_sheet
            gr.update(choices=new_choices, value=first),          # adv_sheet_radio
            gr.update(choices=cols, value=t_val),                  # adv_title_col
            render_adv_table(first, adv_config) if first else "", # adv_table_html
            invalidate(),                                          # gen_file
        )

    # ── File-load handler ─────────────────────────────────────────

    def _make_load_outputs(ok, msg, sheets):
        """Build the tuple returned by both load handlers."""
        if not ok:
            empty = gr.update(choices=[])
            return (msg, empty, empty, {"groups": [], "sheets": {}}, "",
                    empty, gr.update(choices=[], value=None),
                    "<p style='color:#888;font-size:13px'>No groups yet.</p>",
                    "<p style='color:#888'>Load a file and select a sheet.</p>",
                    gr.update(choices=[]), invalidate(), gr.update(value=True))

        conv.selected_sheets = set(s for s in sheets if s != "structure")
        sel_list     = sorted(conv.selected_sheets)
        first        = sel_list[0] if sel_list else None
        new_root     = len(sel_list) > 1
        conv.use_file_root = new_root

        for s in sel_list:
            if s in conv.sheets_data and not conv.sheet_title_cols.get(s):
                cols = list(conv.sheets_data[s].columns)
                if cols: conv.sheet_title_cols[s] = cols[0]

        adv_config = _init_adv_config(sel_list)
        first_cols  = list(conv.sheets_data[first].columns) if first and first in conv.sheets_data else []
        t_val       = adv_config.get("sheets", {}).get(first, {}).get("title_col",
                          first_cols[0] if first_cols else None)
        groups      = adv_config.get("groups", [])

        return (
            msg,                                                         # status
            gr.update(choices=sheets, value=list(conv.selected_sheets)), # sheets_sel
            gr.update(choices=sel_list, value=first),                   # hier_sheet
            adv_config,                                                   # adv_state
            first or "",                                                  # adv_cur_sheet
            gr.update(choices=sel_list, value=first),                   # adv_sheet_radio
            gr.update(choices=first_cols, value=t_val),                  # adv_title_col
            render_adv_groups(adv_config),                               # adv_grp_display
            render_adv_table(first, adv_config) if first else "",        # adv_table_html
            gr.update(choices=groups),                                    # adv_grp_del_dropdown
            invalidate(),                                                 # gen_file
            gr.update(value=new_root),                                   # use_file_root_cb
        )

    def do_load_file(file):
        if not file: return _make_load_outputs(False, "No file selected.", [])
        ok, msg, ch = conv.load_xlsx(file.name)
        return _make_load_outputs(ok, msg, ch)

    def do_load_gs(url, json_txt):
        ok, msg, ch = conv.load_gsheet(url, json_txt)
        return _make_load_outputs(ok, msg, ch)

    # ── Generate XML ──────────────────────────────────────────────

    def gen_xml(mode_val, use_l1, use_deep, data_delim, adv_config):
        conv.use_prefix_l1  = use_l1
        conv.use_prefix_deep = use_deep
        conv.data_delimiter = None if data_delim == "No Separation" else data_delim
        if mode_val == "Advanced":
            _sync_adv_to_conv(adv_config)
        try:
            path, n, m = conv.generate(mode_val)
            return gr.update(value=path, visible=True), f"✅ Success! {n} nodes, {m} edges generated."
        except Exception as e:
            traceback.print_exc()
            return gr.update(visible=False), f"❌ **Error:** {str(e)}"

    # ── Template handlers ─────────────────────────────────────────

    def tpl_init_save(login):
        if not (login or "").strip():
            return gr.update(visible=False), gr.update(visible=False), gr.update(visible=False, value=False), \
                   "⚠️ Please enter a login before saving."
        if not conv.loaded_sheets:
            return gr.update(visible=False), gr.update(visible=False), gr.update(visible=False, value=False), \
                   "⚠️ Please load a file before saving a template."
        return gr.update(visible=True), gr.update(visible=False), gr.update(visible=False, value=False), ""

    def tpl_confirm_save(login, name, overwrite):
        login = (login or "").strip()
        if not login:
            return gr.update(visible=True), gr.update(visible=False), gr.update(), "⚠️ Login is required."
        if not conv.loaded_sheets:
            return gr.update(visible=False), gr.update(visible=False), gr.update(), "⚠️ No file loaded."
        name = (name or "").strip() or f"template_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"
        all_tpls  = _tmpl_load_all()
        user_tpls = all_tpls.get(login, {})
        if name in user_tpls and not overwrite:
            return (gr.update(visible=True), gr.update(visible=True), gr.update(),
                    f"⚠️ Template **{name}** already exists. Check 'Overwrite if exists' and confirm.")
        word = "overwritten" if (name in user_tpls and overwrite) else "saved"
        user_tpls[name] = {"created": datetime.datetime.now().isoformat(),
                           "structural_metadata": _tmpl_snapshot_structure(),
                           "config": _tmpl_collect_config()}
        all_tpls[login] = user_tpls
        _tmpl_save_all(all_tpls)
        return gr.update(visible=False), gr.update(visible=False), gr.update(value=""), \
               f"✅ Template **{name}** {word} successfully."

    def tpl_init_load(login):
        login = (login or "").strip()
        if not login:
            return gr.update(visible=False), gr.update(visible=False), gr.update(choices=[]), \
                   "⚠️ Please enter a login before loading."
        if not conv.loaded_sheets:
            return gr.update(visible=False), gr.update(visible=False), gr.update(choices=[]), \
                   "⚠️ Please load a file before applying a template."
        all_tpls  = _tmpl_load_all()
        user_tpls = all_tpls.get(login, {})
        if not user_tpls:
            return gr.update(visible=False), gr.update(visible=False), gr.update(choices=[]), \
                   f"ℹ️ No templates found for login **{login}**."
        names = sorted(user_tpls.keys())
        return gr.update(visible=True), gr.update(visible=False), gr.update(choices=names, value=names[0]), ""

    def tpl_apply_fn(login, tpl_name):
        def _noop(msg):
            return (msg, gr.update(), gr.update(), gr.update(), gr.update(), gr.update(),
                    gr.update(), gr.update(), gr.update(), gr.update(), gr.update(),
                    gr.update(), gr.update(), gr.update(), gr.update(visible=False), gr.update())

        login = (login or "").strip()
        if not login or not tpl_name:
            return _noop("⚠️ Please select a template to apply.")
        template = _tmpl_load_all().get(login, {}).get(tpl_name)
        if not template:
            return _noop(f"⚠️ Template **{tpl_name}** not found for login **{login}**.")

        applied, skipped, match_type = _tmpl_apply(template)
        if match_type == "none":
            return _noop(f"❌ Template **{tpl_name}** is incompatible with the current file.")

        parts = []
        if applied: parts.append("**Applied:**\n" + "\n".join(f"- {a}" for a in applied))
        if skipped: parts.append("**Skipped:**\n" + "\n".join(f"- {s}" for s in skipped))
        summary_md = "\n\n".join(parts)
        status_msg = (f"✅ Template **{tpl_name}** applied — {len(applied)} settings applied."
                      if match_type == "full" else
                      f"⚠️ Template **{tpl_name}** partially applied — {len(applied)} applied, {len(skipped)} skipped.")

        sel_list   = sorted(conv.selected_sheets)
        first      = sel_list[0] if sel_list else None
        hier_avail = []
        if first and first in conv.sheets_data:
            all_c  = list(conv.sheets_data[first].columns)
            used_c = [lv["col"] if isinstance(lv, dict) else str(lv)
                      for lv in conv.hierarchy_configs.get(first, [])]
            hier_avail = [c for c in all_c if c not in used_c]

        adv_config = _rebuild_adv_from_conv(sel_list)
        first_cols = list(conv.sheets_data[first].columns) if first and first in conv.sheets_data else []
        t_val      = adv_config.get("sheets", {}).get(first, {}).get("title_col",
                         first_cols[0] if first_cols else None)
        delim_val  = conv.data_delimiter if conv.data_delimiter else "No Separation"

        return (
            status_msg,                                             # tpl_status
            gr.update(visible=False),                               # tpl_load_panel
            gr.update(value=sel_list),                              # sheets_sel
            gr.update(value=first),                                 # hier_sheet
            format_hier_df(first),                                  # hier_data
            gr.update(choices=hier_avail),                          # hier_add_col
            gr.update(value=conv.edge_direction),                   # direction
            gr.update(value=conv.use_file_root),                    # use_file_root_cb
            gr.update(value=conv.exclude_title_from_data),          # exclude_title_cb
            gr.update(value=delim_val),                             # global_data_delim
            adv_config,                                             # adv_state
            gr.update(choices=sel_list, value=first),               # adv_sheet_radio
            render_adv_groups(adv_config),                          # adv_grp_display
            render_adv_table(first, adv_config) if first else "",   # adv_table_html
            gr.update(visible=bool(summary_md)),                    # tpl_summary_accordion
            summary_md,                                             # tpl_summary
        )

    # ── BINDINGS ─────────────────────────────────────────────────

    mode_toggle.change(toggle_mode, inputs=[mode_toggle],
                       outputs=[step2_simple_row, advanced_group])

    use_file_root_cb.change(set_root_flag,    inputs=[use_file_root_cb], outputs=[gen_file])
    exclude_title_cb.change(set_exclude_title, inputs=[exclude_title_cb], outputs=[gen_file])
    direction.change(set_direction,            inputs=[direction],        outputs=[gen_file])
    global_data_delim.change(invalidate,                                  outputs=[gen_file])
    hier_show_class_l1.change(set_l1,          inputs=[hier_show_class_l1], outputs=[gen_file])
    hier_show_class_deep.change(set_deep,      inputs=[hier_show_class_deep], outputs=[gen_file])

    # Common outputs for file-load events
    _load_outs = [status, sheets_sel, hier_sheet, adv_state, adv_cur_sheet,
                  adv_sheet_radio, adv_title_col, adv_grp_display, adv_table_html,
                  adv_grp_del_dropdown, gen_file, use_file_root_cb]
    load_file_btn.click(do_load_file, inputs=[up_file],          outputs=_load_outs)
    load_gs_btn.click(  do_load_gs,  inputs=[gs_url, gs_json],   outputs=_load_outs)

    # Simple-mode hierarchy
    hier_sheet.change(on_hier_sheet_change, inputs=[hier_sheet], outputs=[hier_add_col, hier_data])
    hier_add_btn.click(add_hier_level,      inputs=[hier_sheet, hier_add_col, hier_delim],
                       outputs=[hier_data, hier_add_col, gen_file])
    hier_data.select(on_hier_select, None, hier_sel)
    hier_up_btn.click(  lambda s, i: hier_move(s, i, "up"),   inputs=[hier_sheet, hier_sel],
                        outputs=[hier_data, gen_file])
    hier_down_btn.click(lambda s, i: hier_move(s, i, "down"), inputs=[hier_sheet, hier_sel],
                        outputs=[hier_data, gen_file])
    hier_del_btn.click( hier_delete, inputs=[hier_sheet, hier_sel],
                        outputs=[hier_data, hier_add_col, gen_file])

    # Advanced mode — included sheets
    sheets_sel.change(on_adv_sheets_sel,
                      inputs=[sheets_sel, adv_state],
                      outputs=[hier_sheet, adv_state, adv_cur_sheet,
                               adv_sheet_radio, adv_title_col, adv_table_html, gen_file])

    # Advanced mode — sheet tab switch
    adv_sheet_radio.change(on_adv_sheet_change,
                           inputs=[adv_sheet_radio, adv_state],
                           outputs=[adv_cur_sheet, adv_num_levels, adv_title_col, adv_table_html])

    # Advanced mode — hierarchy levels
    adv_num_levels.change(on_adv_num_levels,
                          inputs=[adv_num_levels, adv_cur_sheet, adv_state],
                          outputs=[adv_state, adv_table_html])

    # Advanced mode — title column
    adv_title_col.change(on_adv_title_col,
                         inputs=[adv_title_col, adv_cur_sheet, adv_state],
                         outputs=[adv_state])

    # Advanced mode — groups
    adv_grp_add_btn.click(on_adv_add_group,
                          inputs=[adv_grp_input, adv_cur_sheet, adv_state],
                          outputs=[adv_state, adv_grp_input, adv_grp_display,
                                   adv_grp_del_dropdown, adv_table_html])
    adv_grp_del_btn.click(on_adv_del_group,
                          inputs=[adv_grp_del_dropdown, adv_cur_sheet, adv_state],
                          outputs=[adv_state, adv_grp_display, adv_grp_del_dropdown, adv_table_html])

    # Advanced mode — JS table changes
    adv_col_change.change(on_adv_col_change,
                          inputs=[adv_col_change, adv_state],
                          outputs=[adv_state])

    # Generate
    gen_btn.click(gen_xml,
                  inputs=[mode_toggle, hier_show_class_l1, hier_show_class_deep,
                          global_data_delim, adv_state],
                  outputs=[gen_file, gen_status])

    # ── Template bindings ─────────────────────────────────────────
    tpl_save_btn.click(tpl_init_save, inputs=[tpl_login],
                       outputs=[tpl_save_panel, tpl_load_panel, tpl_overwrite_cb, tpl_status])
    tpl_confirm_save_btn.click(tpl_confirm_save,
                               inputs=[tpl_login, tpl_name_input, tpl_overwrite_cb],
                               outputs=[tpl_save_panel, tpl_overwrite_cb, tpl_name_input, tpl_status])
    tpl_load_btn.click(tpl_init_load, inputs=[tpl_login],
                       outputs=[tpl_load_panel, tpl_save_panel, tpl_list, tpl_status])
    tpl_confirm_load_btn.click(
        tpl_apply_fn,
        inputs=[tpl_login, tpl_list],
        outputs=[
            tpl_status, tpl_load_panel,
            sheets_sel, hier_sheet, hier_data, hier_add_col,
            direction, use_file_root_cb, exclude_title_cb, global_data_delim,
            adv_state, adv_sheet_radio, adv_grp_display, adv_table_html,
            tpl_summary_accordion, tpl_summary,
        ],
    )

    app.load(None, None, gs_json, js="() => localStorage.getItem('gs_creds') || ''")
    gs_json.change(None, gs_json, None, js="(v) => localStorage.setItem('gs_creds', v)")
 
if __name__ == "__main__":
    app.launch()