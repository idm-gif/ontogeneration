#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Sheets → XML Graph Converter
Fixed:
1. Added "Data Separator" to split cell values into multiple data entries (Array support).
2. Applies globally to all data attributes in both Simple and Advanced modes.
3. Added interactive Column Configuration Table inside Step 1 for Advanced Mode.
4. Enforces user-selected column types (image, youtube, link).
5. 'Auto' detects standard URLs as 'link' (not 'text') and cleans up media content strings.
6. TRANSFERRED HIERARCHY: Explicit level-count selection per sheet.
7. TRANSFERRED GROUPING: Added dynamic cross-sheet grouping directly into the Advanced table.
8. MOVED SHEET SELECTION: Brought "Include Sheets" to the top of Step 1 for better UX.
9. RESTORED GLOBAL PREFIXES: Removed per-column prefixing and restored global "Show Column Name
   in Level 1/Deeper Levels" checkboxes, moving them to Step 1 so they apply universally.
10. FIXED ADVANCED MODE: Replaced JavaScript/hidden-element hack with native Gradio components
    so that per-column type, hierarchy, delimiter, and grouping controls actually work.
"""
import os
import re
import uuid
import json
import traceback
import typing as t
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

MAX_COLUMNS = 30  # Maximum columns rendered in Advanced config


class GraphNode:
    def __init__(self, name: str, data_elements: t.List[dict] = None,
                 node_type: str = "", is_intermediate: bool = False):
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
        self.current_mode: str = "Typical (Simple)"

        # Simple Mode Storage
        self.hierarchy_configs: dict[str, list[dict]] = {}
        self.column_type_preferences: dict[str, dict[str, str]] = {}

        # Advanced Mode Storage
        self.adv_col_config: dict[str, dict[str, dict]] = {}
        self.adv_sheet_levels: dict[str, int] = {}
        self.adv_grouping_enabled: bool = False
        self.adv_groups: list[str] = []

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
        self.data_delimiter: str = ";"

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
            valid_cols = sum(1 for val in row if pd.notna(val) and str(val).strip() != "")
            if valid_cols > max_valid_cols:
                max_valid_cols = valid_cols
                best_idx = i
        return best_idx

    def _reset_state(self):
        self.hierarchy_configs = {}
        self.grouping_rules = []
        self.column_type_preferences = {}
        self.sheet_title_cols = {}
        self.adv_col_config = {}
        self.adv_sheet_levels = {}
        self.adv_grouping_enabled = False
        self.adv_groups = []
        for sheet, df in self.sheets_data.items():
            self.adv_sheet_levels[sheet] = 0
            self.adv_col_config[sheet] = {}
            for c in df.columns:
                self.adv_col_config[sheet][c] = {
                    'type': 'auto', 'hier': 0, 'delim': ';', 'group': 'None'
                }

    def load_xlsx(self, file_path: str) -> tuple[bool, str, list[str]]:
        try:
            xl = pd.ExcelFile(file_path, engine="openpyxl")
            self.sheets_data, self.loaded_sheets = {}, []
            for sheet in xl.sheet_names:
                df_raw = pd.read_excel(xl, sheet_name=sheet, header=None, nrows=25)
                if df_raw.empty:
                    continue
                header_idx = self._find_header_row(df_raw)
                df = pd.read_excel(xl, sheet_name=sheet, header=header_idx)
                df.columns = self._clean_headers(df.columns)
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
            return True, f"Loaded {len(self.loaded_sheets)} sheets.", [
                s for s in self.loaded_sheets if s != "structure"
            ]
        except Exception as e:
            return False, f"Excel Error: {str(e)}", []

    def load_gsheet(self, url: str, creds_json: str) -> tuple[bool, str, list[str]]:
        if not HAS_SHEETS:
            return False, "Error: 'gspread' library not installed.", []
        try:
            if not url or not creds_json:
                return False, "Please provide URL and JSON.", []
            creds_dict = json.loads(creds_json)
            scope = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive',
            ]
            creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
            client = gspread.authorize(creds)
            sheet = client.open_by_url(url)
            self.root_name = re.sub(r'[\\/*?:"<>|]', "", sheet.title)
            self.sheets_data, self.loaded_sheets = {}, []
            for ws in sheet.worksheets():
                data = ws.get_all_values()
                if not data:
                    continue
                df_raw = pd.DataFrame(data)
                header_idx = self._find_header_row(df_raw)
                if header_idx + 1 < len(data):
                    headers = data[header_idx]
                    rows = data[header_idx + 1:]
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
            return True, f"Loaded {len(self.loaded_sheets)} sheets.", [
                s for s in self.loaded_sheets if s != "structure"
            ]
        except Exception as e:
            return False, f"Error: {str(e)}", []

    @staticmethod
    def _extract_url_from_formula(cell_value: str) -> tuple[str, str, bool]:
        if not isinstance(cell_value, str):
            return (str(cell_value), "", False)
        m = re.search(
            r'=HYPERLINK\s*\(\s*"([^"]+)"\s*,\s*"([^"]*)"\s*\)',
            cell_value, re.IGNORECASE,
        )
        if m:
            return (m.group(2) or m.group(1), m.group(1), True)
        if re.match(r'^\s*https?://[^\s]+', cell_value):
            cleaned_url = cell_value.strip()
            return (cleaned_url, cleaned_url, True)
        return (cell_value, "", False)

    @staticmethod
    def _is_image(url: str) -> bool:
        if not url or not isinstance(url, str):
            return False
        return any(
            url.lower().strip().endswith(e)
            for e in [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".svg", ".webp"]
        )

    @staticmethod
    def _is_youtube(url: str) -> bool:
        if not url or not isinstance(url, str):
            return False
        u = url.lower().strip()
        return "youtube.com" in u or "youtu.be" in u

    def _detect_type(self, url_or_text: str) -> str:
        if self._is_image(url_or_text):
            return "image"
        if self._is_youtube(url_or_text):
            return "youtube"
        return "link"

    def _row_to_elements(self, row: pd.Series, cols: list[str],
                         sheet: str) -> list[dict]:
        result = []
        for col in cols:
            val = row.get(col)
            if pd.notna(val) and str(val).strip():
                v_str = str(val).strip()
                values = [v_str]
                if self.data_delimiter and self.data_delimiter != "No Separation":
                    values = [x.strip() for x in v_str.split(self.data_delimiter) if x.strip()]
                for v in values:
                    text, url, has_link = self._extract_url_from_formula(v)
                    pref = self.column_type_preferences.get(sheet, {}).get(col, "auto")
                    if pref == "text":
                        result.append({"tclass": str(col), "link": "", "type": "text", "content": v})
                    else:
                        if pref in ["image", "youtube", "link"] and not has_link:
                            url = v
                            text = v if pref == "link" else ""
                            has_link = True
                        if has_link:
                            ltype = self._detect_type(url) if pref == "auto" else pref
                            if ltype in ["image", "youtube"]:
                                text = ""
                            elif not text:
                                text = url
                            result.append({"tclass": str(col), "link": url, "type": ltype, "content": text})
                        else:
                            ltype = "text" if pref == "auto" else pref
                            result.append({"tclass": str(col), "link": "", "type": ltype, "content": v})
        return result

    def _create_edge(self, parent: str, child: str) -> GraphEdge:
        if self.edge_direction == "left":
            return GraphEdge(parent, child)
        return GraphEdge(child, parent)

    def _get_or_create_row_node(self, sheet_name: str, row: pd.Series,
                                cols: list[str]) -> GraphNode:
        node_id = f"{sheet_name}_{row.name}"
        if node_id in self.node_map:
            return self.node_map[node_id]
        title = None
        configured_col = self.sheet_title_cols.get(sheet_name)
        if configured_col and configured_col in cols:
            val = row.get(configured_col)
            if pd.notna(val) and str(val).strip():
                title = str(val).strip()
        if not title and self.node_name_strategy == "Name":
            name_col = next((c for c in cols if c.lower() == "name"), None)
            if name_col:
                val = row.get(name_col)
                if pd.notna(val) and str(val).strip():
                    title = str(val).strip()
        if not title and len(cols) > 0:
            val = row.iloc[0]
            if pd.notna(val) and str(val).strip():
                title = str(val).strip()
        if not title:
            for c in cols:
                val = row.get(c)
                if pd.notna(val) and str(val).strip() and not str(val).isnumeric():
                    title = str(val).strip()
                    break
        if not title:
            title = f"{sheet_name}_{row.name}"
        if title in self.node_map:
            return self.node_map[title]
        node = GraphNode(title, self._row_to_elements(row, cols, sheet_name), sheet_name)
        self.nodes.append(node)
        self.node_map[title] = node
        self.node_map[node_id] = node
        return node

    def _enrich_phantom_data(self, col: str, val: str, ctx: tuple) -> list[dict]:
        data_elements = []
        if col not in self.sheets_data:
            return data_elements
        lookup_df = self.sheets_data[col]
        lookup_cols = list(lookup_df.columns)
        matching_rows = pd.DataFrame()
        if len(lookup_df.columns) > 0:
            title_col = lookup_df.columns[0]
            matching_rows = lookup_df[
                lookup_df[title_col].astype(str).str.strip() == str(val).strip()
            ]
        if not matching_rows.empty:
            match_row = matching_rows.iloc[0]
            for lookup_col in lookup_cols:
                elem_val = match_row.get(lookup_col)
                if pd.notna(elem_val) and str(elem_val).strip():
                    v = str(elem_val)
                    text, url, has_link = self._extract_url_from_formula(v)
                    pref = self.column_type_preferences.get(col, {}).get(lookup_col, "auto")
                    if pref == "text":
                        data_elements.append({"tclass": str(lookup_col), "link": "", "type": "text", "content": v})
                    else:
                        if pref in ["image", "youtube", "link"] and not has_link:
                            url = v
                            text = v if pref == "link" else ""
                            has_link = True
                        if has_link:
                            ltype = self._detect_type(url) if pref == "auto" else pref
                            if ltype in ["image", "youtube"]:
                                text = ""
                            elif not text:
                                text = url
                            data_elements.append({"tclass": str(lookup_col), "link": url, "type": ltype, "content": text})
                        else:
                            ltype = "text" if pref == "auto" else pref
                            data_elements.append({"tclass": str(lookup_col), "link": "", "type": ltype, "content": v})
        return data_elements

    def _get_phantom(self, ctx: tuple, col: str, val: str) -> str:
        key = (ctx, col, val)
        if key in self.phantom_nodes:
            return self.phantom_nodes[key]
        is_level_1 = (len(ctx) == 0)
        use_prefix = self.use_prefix_l1 if is_level_1 else self.use_prefix_deep
        name = f"{col}: {val}" if use_prefix else f"{val}"
        data_elements = self._enrich_phantom_data(col, val, ctx)
        if name in self.node_map:
            self.phantom_nodes[key] = name
            return name
        node = GraphNode(name, data_elements, col)
        self.nodes.append(node)
        self.node_map[name] = node
        self.phantom_nodes[key] = name
        return node.name

    def _split_multi_value(self, value: str, delimiter: str = None) -> list[str]:
        if not delimiter or delimiter == "No Separation":
            return [value]
        return [v.strip() for v in str(value).split(delimiter) if v.strip()]

    def _apply_mode4_rule(self, rule: dict):
        if "pairs" not in rule:
            return
        group_name = rule.get("name", "Group")
        default_delimiter = rule.get("delimiter", self.multi_sub_delimiter)
        fields = []
        for item in rule["pairs"]:
            if len(item) == 3:
                s, f, delim = item
            else:
                s, f = item
                delim = default_delimiter
            if s in self.sheets_data and s in self.selected_sheets:
                df = self.sheets_data[s]
                if f in df.columns:
                    vals = set(str(v).strip() for v in df[f].dropna() if str(v).strip())
                    fields.append((s, f, vals, delim))
        if not fields:
            return
        for s, _, _, _ in fields:
            self.sheets_in_grouping.add(s)

        value_to_sf = {}
        for s, f, _, delim in fields:
            df = self.sheets_data[s]
            for val in df[f].dropna():
                v_str = str(val).strip()
                if not v_str:
                    continue
                sub_values = self._split_multi_value(v_str, delim)
                for sub_val in sub_values:
                    value_to_sf.setdefault(sub_val, []).append((s, f, val))

        for value, sheet_field_originals in value_to_sf.items():
            v_key = f"M4_{value}"
            data_elements = self._enrich_phantom_data(group_name, value, ())
            if v_key not in self.node_map:
                v_node = GraphNode(value, data_elements, group_name, is_intermediate=True)
                self.nodes.append(v_node)
                self.node_map[v_key] = v_node
                self.node_map[value] = v_node
            self.edges.append(self._create_edge(self.root_name, value))

            processed_combos = set()
            for s, f, orig_val in sheet_field_originals:
                combo_key = (s, f, orig_val)
                if combo_key in processed_combos:
                    continue
                processed_combos.add(combo_key)
                root_node_name = s
                if root_node_name not in self.node_map:
                    root_node = GraphNode(root_node_name, [], s)
                    self.nodes.append(root_node)
                    self.node_map[root_node_name] = root_node
                self.edges.append(self._create_edge(value, root_node_name))
                matches = self.sheets_data[s][
                    self.sheets_data[s][f].astype(str) == str(orig_val)
                ]
                for _, row in matches.iterrows():
                    node = self._get_or_create_row_node(s, row, list(self.sheets_data[s].columns))
                    self.edges.append(self._create_edge(root_node_name, node.name))

        if self.show_empty_nodes:
            empty_sheet_rows = []
            for s, f, _, _ in fields:
                df_s = self.sheets_data[s]
                empty_mask = df_s[f].isna() | (df_s[f].astype(str).str.strip() == "")
                empty_rows = df_s[empty_mask]
                if not empty_rows.empty:
                    empty_sheet_rows.append((s, f, empty_rows))
            if empty_sheet_rows:
                empty_key = "(empty)"
                if empty_key not in self.node_map:
                    empty_node = GraphNode("empty", [], group_name, is_intermediate=True)
                    self.nodes.append(empty_node)
                    self.node_map[empty_key] = empty_node
                    self.node_map["empty"] = empty_node
                self.edges.append(self._create_edge(self.root_name, "empty"))
                for s, f, empty_rows in empty_sheet_rows:
                    if s not in self.node_map:
                        s_node = GraphNode(s, [], s)
                        self.nodes.append(s_node)
                        self.node_map[s] = s_node
                    self.edges.append(self._create_edge("empty", s))
                    for _, row in empty_rows.iterrows():
                        node = self._get_or_create_row_node(s, row, list(self.sheets_data[s].columns))
                        self.edges.append(self._create_edge(s, node.name))

    def apply_groupings(self):
        self.sheets_in_grouping = set()
        for rule in self.grouping_rules:
            self._apply_mode4_rule(rule)

    def _build_hier(self, df: pd.DataFrame, cfg: list[dict], parent: str,
                    sheet: str, lvl: int = 0, current_filter: dict = None):
        if current_filter is None:
            current_filter = {}
        if lvl >= len(cfg):
            filtered_rows = df.copy()
            for col, val in current_filter.items():
                if val is None:
                    filtered_rows = filtered_rows[
                        filtered_rows[col].isna() | (filtered_rows[col].astype(str).str.strip() == "")
                    ]
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
        if col not in df.columns:
            return

        curr = df.copy()
        for col_name, val in current_filter.items():
            if val is None:
                curr = curr[curr[col_name].isna() | (curr[col_name].astype(str).str.strip() == "")]
            else:
                curr = curr[curr[col_name].astype(str) == str(val)]

        processed_vals = set()
        for index, row in curr.iterrows():
            val = row.get(col)
            if pd.isna(val) or not str(val).strip():
                continue
            val_str = str(val).strip()
            if val_str in processed_vals:
                continue
            processed_vals.add(val_str)
            sub_values = self._split_multi_value(val_str, level_delim)
            for sub_val in sub_values:
                phantom_name = self._get_phantom(
                    tuple(sorted([(c, str(v)) for c, v in current_filter.items()])),
                    col, sub_val,
                )
                edge_exists = any(
                    e.node1 == parent and e.node2 == phantom_name for e in self.edges
                )
                if not edge_exists:
                    self.edges.append(self._create_edge(parent, phantom_name))
                nf = dict(current_filter)
                nf[col] = val_str
                self._build_hier(df, cfg, phantom_name, sheet, lvl + 1, nf)

        if self.show_empty_nodes:
            empty_mask = curr[col].isna() | (curr[col].astype(str).str.strip() == "")
            if empty_mask.any():
                empty_val = "(empty)"
                empty_name = self._get_phantom(
                    tuple(sorted([(c, str(v)) for c, v in current_filter.items()])),
                    col, empty_val,
                )
                edge_exists = any(
                    e.node1 == parent and e.node2 == empty_name for e in self.edges
                )
                if not edge_exists:
                    self.edges.append(self._create_edge(parent, empty_name))
                nf = dict(current_filter)
                nf[col] = None
                self._build_hier(df, cfg, empty_name, sheet, lvl + 1, nf)

    def generate(self, mode_val: str) -> tuple[str, int, int]:
        self.current_mode = mode_val
        if not self.selected_sheets and not self.loaded_sheets:
            raise RuntimeError("No sheets loaded")

        if mode_val == "Typical (Simple)":
            self.edge_direction = "right"
            self.node_name_strategy = "first"
            if not self.selected_sheets:
                self.selected_sheets = set(s for s in self.loaded_sheets if s != "structure")

        elif mode_val == "Advanced":
            self.hierarchy_configs = {}
            self.column_type_preferences = {}
            self.grouping_rules = []

            # Map Groupings from adv_col_config
            if self.adv_grouping_enabled and self.adv_groups:
                for g_name in self.adv_groups:
                    pairs = []
                    for sheet, c_cfg in self.adv_col_config.items():
                        if sheet not in self.selected_sheets:
                            continue
                        for col, cfg in c_cfg.items():
                            if cfg.get('group') == g_name:
                                pairs.append((sheet, col, cfg.get('delim', ';')))
                    if pairs:
                        self.grouping_rules.append({"name": g_name, "pairs": pairs})

            # Map Types & Hierarchies from adv_col_config
            for sheet, c_cfg in self.adv_col_config.items():
                if sheet not in self.selected_sheets:
                    continue
                self.column_type_preferences[sheet] = {}
                for col, cfg in c_cfg.items():
                    self.column_type_preferences[sheet][col] = cfg.get('type', 'auto')
                hier_cols = [
                    (col, cfg.get('hier', 0), cfg.get('delim', ';'))
                    for col, cfg in c_cfg.items()
                    if cfg.get('hier', 0) > 0
                ]
                hier_cols.sort(key=lambda x: x[1])
                if hier_cols:
                    self.hierarchy_configs[sheet] = [
                        {'col': hc[0], 'delim': None if hc[2] == 'No Separation' else hc[2]}
                        for hc in hier_cols
                    ]

        self.nodes, self.edges, self.node_map, self.phantom_nodes = [], [], {}, {}

        if self.use_file_root:
            root = GraphNode(self.root_name)
            self.nodes.append(root)
            self.node_map[self.root_name] = root

        self.apply_groupings()

        for name in sorted(self.selected_sheets):
            if name not in self.sheets_data or name in self.sheets_in_grouping:
                continue
            if name not in self.node_map:
                n = GraphNode(name, [], name)
                self.nodes.append(n)
                self.node_map[name] = n
            if self.use_file_root:
                self.edges.append(self._create_edge(self.root_name, name))
            df = self.sheets_data[name]
            cfg = self.hierarchy_configs.get(name, [])
            if cfg:
                self._build_hier(df, cfg, name, name, 0)
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
            if (i + 1) % 5 == 0:
                x, y = 0, y + 150

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
        xml_str = minidom.parseString(
            ET.tostring(graph, encoding="unicode")
        ).toprettyxml(indent="  ")
        with open(out_file, "w", encoding="utf-8") as f:
            f.write(xml_str)
        return out_file, len(self.nodes), len(self.edges)


# --- [UI] ---

css = """
:root { --primary: #2563eb; --bg: #f3f4f6; --card-bg: #ffffff; }
body, .gradio-container { background: var(--bg); font-family: 'Inter', sans-serif; }
.step-card { background: var(--card-bg); border: 1px solid #e5e7eb; border-radius: 12px;
             padding: 20px; margin-bottom: 16px; box-shadow: 0 2px 4px rgba(0,0,0,0.02); }
.step-header { font-size: 16px; font-weight: 700; color: #1f2937; margin-bottom: 16px;
               padding-bottom: 8px; border-bottom: 2px solid #f3f4f6;
               text-transform: uppercase; letter-spacing: 0.05em; }
.clean-row { gap: 16px; align-items: flex-end; margin-bottom: 8px; }
.btn-main { background-color: #2563eb !important; color: white !important;
            font-size: 15px; font-weight: 600; height: 48px; }
.btn-action { background-color: #ffffff !important; color: #374151 !important;
              border: 1px solid #d1d5db !important; height: 40px; font-size: 13px; }
.btn-del { background-color: #fee2e2 !important; color: #b91c1c !important;
           border: 1px solid #fca5a5 !important; height: 40px; font-size: 13px; }
input, select, textarea { font-size: 14px !important; border-radius: 6px !important; }
textarea { height: 80px !important; }
.col-config-label { font-weight: 600; color: #374151; padding: 6px 0; min-width: 120px;
                    overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
"""

conv = Converter()

# ── Helper: build a serialisable snapshot of the adv config for a sheet ──

def _adv_config_state_for_sheet(sheet: str) -> dict:
    """Return the current adv_col_config for *sheet* as a plain dict."""
    return conv.adv_col_config.get(sheet, {})


def _apply_adv_change(sheet: str, col: str, field: str, value):
    """Mutate conv.adv_col_config in-place for a single field change."""
    if sheet not in conv.adv_col_config or col not in conv.adv_col_config[sheet]:
        return
    if field == "hier":
        new_lvl = int(value)
        # Ensure no two columns share the same hierarchy level
        if new_lvl > 0:
            for c_name, c_cfg in conv.adv_col_config[sheet].items():
                if c_name != col and c_cfg.get("hier") == new_lvl:
                    c_cfg["hier"] = 0
        conv.adv_col_config[sheet][col]["hier"] = new_lvl
    else:
        conv.adv_col_config[sheet][col][field] = value


# ── Dynamic column-config component builder ──

def _build_col_config_components(sheet: str, grouping_enabled: bool,
                                 group_names: list[str], max_levels: int):
    """
    Return a list of Gradio component rows for every column of *sheet*.
    Each row has: Label | Hierarchy dropdown | Delimiter dropdown |
                  (optional) Grouping dropdown | Type dropdown
    """
    if not sheet or sheet not in conv.sheets_data:
        return []
    cols = list(conv.sheets_data[sheet].columns)[:MAX_COLUMNS]
    config = conv.adv_col_config.get(sheet, {})

    hier_choices = ["None"] + [f"Level {i}" for i in range(1, max_levels + 1)]
    group_choices = ["None"] + list(group_names)

    rows = []
    for c in cols:
        cfg = config.get(c, {"type": "auto", "hier": 0, "delim": ";", "group": "None"})
        h_val = f"Level {cfg['hier']}" if cfg.get("hier", 0) > 0 else "None"
        row_components = {
            "col_name": c,
            "hier_value": h_val,
            "hier_choices": hier_choices,
            "delim_value": cfg.get("delim", ";"),
            "group_value": cfg.get("group", "None"),
            "group_choices": group_choices,
            "type_value": cfg.get("type", "auto"),
            "show_group": grouping_enabled,
        }
        rows.append(row_components)
    return rows


with gr.Blocks(css=css, title="Graph Converter") as app:

    # ── State ──
    adv_sheet_state = gr.State("")   # currently selected sheet in Advanced tab

    # --- STEP 1 ---
    with gr.Group(elem_classes=["step-card"]):
        gr.HTML('<div class="step-header">Step 1: Setup & Source</div>')
        with gr.Tabs():
            with gr.Tab("Upload Excel"):
                with gr.Row(elem_classes=["clean-row"]):
                    up_file = gr.File(label="Excel File (.xlsx)", file_count="single",
                                      file_types=[".xlsx"])
                    load_file_btn = gr.Button("Load Excel", variant="primary",
                                              elem_classes=["btn-main"])
            with gr.Tab("Google Sheets"):
                gr.Markdown("Enter URL and Service Account JSON (saved in your browser).")
                gs_url = gr.Textbox(label="Google Sheet URL")
                gs_json = gr.Textbox(label="Service Account JSON", lines=3, elem_id="gs_creds")
                load_gs_btn = gr.Button("Load Google Sheet", variant="primary",
                                        elem_classes=["btn-main"])

        gr.HTML('<div style="height:16px"></div>')

        # Global Sheet Selection
        gr.Markdown("### Included Sheets")
        sheets_sel = gr.CheckboxGroup(choices=[], label="Select which sheets to process")

        gr.HTML('<div style="height:16px"></div>')

        with gr.Row(elem_classes=["clean-row"]):
            mode_toggle = gr.Radio(
                choices=["Typical (Simple)", "Advanced"], value="Typical (Simple)",
                label="Mode", info="Advanced mode unlocks grouping and custom types.",
            )
            use_file_root_cb = gr.Checkbox(label="Use File Name as Root Node", value=True)
            global_data_delim = gr.Dropdown(choices=DELIMITER_CHOICES, value=";",
                                            label="Data Field Separator")

        with gr.Row(elem_classes=["clean-row"]):
            hier_show_class_l1 = gr.Checkbox(
                label="Show Column Name in Level 1 (e.g., Column: Value)", value=False)
            hier_show_class_deep = gr.Checkbox(
                label="Show Column Name in Deeper Levels", value=False)

        status = gr.Markdown(value="Ready to load data.")

        # ────────────────────────────────────────────────────────────
        # ADVANCED COLUMN CONFIGURATION  (native Gradio components)
        # ────────────────────────────────────────────────────────────
        with gr.Column(visible=False) as advanced_step1_col:
            gr.HTML('<div style="height:24px"></div>')
            gr.HTML('<div class="step-header" style="color:#7c3aed; border-color:#ddd6fe;">'
                    'Advanced: Column Configuration</div>')
            gr.Markdown("Configure hierarchy, delimiters, grouping, and media types "
                        "individually per column.")

            adv_group_enable_cb = gr.Checkbox(label="Enable Cross-Sheet Grouping", value=False)
            with gr.Column(visible=False) as adv_group_manager:
                with gr.Row(elem_classes=["clean-row"]):
                    adv_new_group_name = gr.Textbox(label="New Group Name", scale=2)
                    adv_add_group_btn = gr.Button("Add Group", elem_classes=["btn-action"],
                                                  scale=1)
                    adv_del_group_dropdown = gr.Dropdown(choices=[], label="Select to Remove",
                                                        scale=2)
                    adv_del_group_btn = gr.Button("Remove", elem_classes=["btn-del"], scale=1)

            gr.HTML('<div style="height:12px"></div>')

            with gr.Row(elem_classes=["clean-row"]):
                types_tab_sheet = gr.Radio(choices=[], label="Select Sheet Configuration",
                                           interactive=True, scale=2)
                hier_level_count = gr.Dropdown(
                    choices=[str(i) for i in range(11)], value="0",
                    label="Number of Hierarchy Levels", interactive=True, scale=1,
                )

            # Apply Config button — reads the dropdowns and writes to conv state
            adv_apply_btn = gr.Button("Apply Column Settings", elem_classes=["btn-action"])
            adv_status = gr.Markdown("")

            # ── Per-column config area ──
            # We create MAX_COLUMNS rows, each with label + dropdowns.
            # Rows are shown/hidden based on actual column count.
            col_config_rows = []  # list of dicts with gr components
            for idx in range(MAX_COLUMNS):
                with gr.Row(visible=False, elem_classes=["clean-row"]) as crow:
                    lbl = gr.Markdown("", elem_classes=["col-config-label"])
                    dd_hier = gr.Dropdown(
                        choices=["None"], value="None", label="Hierarchy",
                        interactive=True, scale=1, min_width=100,
                    )
                    dd_delim = gr.Dropdown(
                        choices=DELIMITER_CHOICES, value=";", label="Separator",
                        interactive=True, scale=1, min_width=80,
                    )
                    dd_group = gr.Dropdown(
                        choices=["None"], value="None", label="Grouping",
                        interactive=True, scale=1, min_width=100, visible=False,
                    )
                    dd_type = gr.Dropdown(
                        choices=TYPE_CHOICES, value="auto", label="Type",
                        interactive=True, scale=1, min_width=100,
                    )
                col_config_rows.append({
                    "row": crow,
                    "label": lbl,
                    "hier": dd_hier,
                    "delim": dd_delim,
                    "group": dd_group,
                    "type": dd_type,
                })

    # --- STEP 2 ---
    with gr.Row():
        # Simple Mode Hierarchy
        with gr.Column(scale=3) as simple_hier_group:
            with gr.Group(elem_classes=["step-card"]):
                gr.HTML('<div class="step-header" style="color:#2563eb; border-color:#bfdbfe;">'
                        'Step 2: Hierarchy</div>')
                with gr.Row(elem_classes=["clean-row"]):
                    hier_sheet = gr.Dropdown(choices=[], label="Select Entity (Sheet)", scale=2)
                    hier_add_col = gr.Dropdown(choices=[],
                                               label="Select Column to Drill Down", scale=2)
                    hier_delim = gr.Dropdown(choices=DELIMITER_CHOICES, value=";",
                                             label="Separator", scale=1)
                    hier_add_btn = gr.Button("+ Add Level", elem_classes=["btn-action"], scale=1)
                hier_data = gr.Dataframe(
                    headers=["Level", "Column Name", "Separator"],
                    datatype=["number", "str", "str"], row_count=0,
                    col_count=(3, "fixed"), interactive=False, label="Structure Preview",
                )
                with gr.Row(elem_classes=["clean-row"]):
                    hier_up_btn = gr.Button("Move Up", elem_classes=["btn-action"])
                    hier_down_btn = gr.Button("Move Down", elem_classes=["btn-action"])
                    hier_del_btn = gr.Button("Delete Selected", elem_classes=["btn-del"])

        # Advanced (rest of Step 2)
        with gr.Column(scale=2, visible=False) as advanced_group:
            with gr.Group(elem_classes=["step-card"]):
                gr.HTML('<div class="step-header" style="color:#7c3aed; border-color:#ddd6fe;">'
                        'Advanced Configuration</div>')
                gr.Markdown("### Node Title Configuration")
                with gr.Row(elem_classes=["clean-row"]):
                    nt_sheet = gr.Dropdown(choices=[], label="Sheet", scale=2)
                    nt_col = gr.Dropdown(choices=[], label="Title Column", scale=2)
                    nt_btn = gr.Button("Set Title", elem_classes=["btn-action"], scale=1)
                nt_view = gr.Dataframe(headers=["Sheet", "Title Column"],
                                       datatype=["str", "str"], interactive=False)
                gr.HTML('<div style="height:12px"></div>')
                direction = gr.Radio(choices=["left", "right"], value="left",
                                     label="Edge Direction")

    # --- STEP 3 ---
    with gr.Group(elem_classes=["step-card"]):
        gr.HTML('<div class="step-header">Step 3: Generate</div>')
        gen_btn = gr.Button("Generate XML Graph", variant="primary", elem_classes=["btn-main"])
        gen_file = gr.File(label="Download Result", visible=False)
        gen_status = gr.Markdown()

    # ───────────────────── LOGIC ─────────────────────

    def invalidate():
        return gr.update(value=None, visible=False)

    def format_hier_df(sheet_name):
        if not sheet_name or sheet_name not in conv.hierarchy_configs:
            return []
        res = []
        for i, item in enumerate(conv.hierarchy_configs[sheet_name]):
            if isinstance(item, dict):
                res.append([i + 1, item['col'], item.get('delim', ';')])
            else:
                res.append([i + 1, str(item), ';'])
        return res

    def format_nt_df():
        return [[s, c] for s, c in conv.sheet_title_cols.items()]

    # ── Populate / refresh the per-column Gradio dropdowns ──

    def _refresh_col_config_outputs(sheet: str):
        """
        Return a flat list of gr.update() calls for every component in
        col_config_rows (row visibility, label, hier, delim, group, type).
        Total: MAX_COLUMNS * 6 outputs.
        Also reads current state from conv.adv_col_config.
        """
        updates = []
        if not sheet or sheet not in conv.sheets_data:
            for _ in range(MAX_COLUMNS):
                updates.extend([
                    gr.update(visible=False),   # row
                    gr.update(value=""),         # label
                    gr.update(value="None", choices=["None"]),  # hier
                    gr.update(value=";"),        # delim
                    gr.update(value="None", choices=["None"], visible=False),  # group
                    gr.update(value="auto"),     # type
                ])
            return updates

        cols = list(conv.sheets_data[sheet].columns)[:MAX_COLUMNS]
        config = conv.adv_col_config.get(sheet, {})
        max_lvl = conv.adv_sheet_levels.get(sheet, 0)
        hier_choices = ["None"] + [f"Level {i}" for i in range(1, max_lvl + 1)]
        group_choices = ["None"] + list(conv.adv_groups)
        show_group = conv.adv_grouping_enabled

        for idx in range(MAX_COLUMNS):
            if idx < len(cols):
                c = cols[idx]
                cfg = config.get(c, {"type": "auto", "hier": 0, "delim": ";", "group": "None"})
                h_val = f"Level {cfg['hier']}" if cfg.get("hier", 0) > 0 else "None"
                updates.extend([
                    gr.update(visible=True),
                    gr.update(value=f"**{c}**"),
                    gr.update(value=h_val, choices=hier_choices),
                    gr.update(value=cfg.get("delim", ";")),
                    gr.update(value=cfg.get("group", "None"), choices=group_choices,
                              visible=show_group),
                    gr.update(value=cfg.get("type", "auto")),
                ])
            else:
                updates.extend([
                    gr.update(visible=False),
                    gr.update(value=""),
                    gr.update(value="None", choices=["None"]),
                    gr.update(value=";"),
                    gr.update(value="None", choices=["None"], visible=False),
                    gr.update(value="auto"),
                ])
        return updates

    # Flat list of outputs for all column-config components
    _col_cfg_outputs = []
    for entry in col_config_rows:
        _col_cfg_outputs.extend([
            entry["row"], entry["label"], entry["hier"],
            entry["delim"], entry["group"], entry["type"],
        ])

    # ── Save dropdown values for a sheet into conv.adv_col_config ──

    def _save_dropdowns_to_config(sheet, dropdown_values):
        """
        Read the current dropdown values and write them into
        conv.adv_col_config[sheet].  dropdown_values is a flat tuple:
        (hier0, delim0, group0, type0, hier1, …) with 4 values per
        MAX_COLUMNS slot.
        """
        if not sheet or sheet not in conv.sheets_data:
            return
        cols = list(conv.sheets_data[sheet].columns)[:MAX_COLUMNS]
        config = conv.adv_col_config.setdefault(sheet, {})

        for idx, c in enumerate(cols):
            base = idx * 4
            if base + 3 >= len(dropdown_values):
                break
            h_raw = dropdown_values[base]      # "None" or "Level N"
            d_raw = dropdown_values[base + 1]  # delimiter
            g_raw = dropdown_values[base + 2]  # group
            t_raw = dropdown_values[base + 3]  # type

            hier_val = 0
            if h_raw and str(h_raw).startswith("Level "):
                try:
                    hier_val = int(str(h_raw).split(" ")[1])
                except (ValueError, IndexError):
                    hier_val = 0

            cfg = config.setdefault(c, {"type": "auto", "hier": 0, "delim": ";", "group": "None"})
            cfg["hier"] = hier_val
            cfg["delim"] = d_raw if d_raw else ";"
            cfg["group"] = g_raw if g_raw else "None"
            cfg["type"] = t_raw if t_raw else "auto"

        # Enforce unique hierarchy levels
        used_levels = {}
        for c in cols:
            lvl = config.get(c, {}).get("hier", 0)
            if lvl > 0:
                if lvl in used_levels:
                    config[used_levels[lvl]]["hier"] = 0
                used_levels[lvl] = c

    def _apply_col_settings(sheet, *dropdown_values):
        """Manual apply button handler."""
        _save_dropdowns_to_config(sheet, dropdown_values)
        if not sheet or sheet not in conv.sheets_data:
            return "No sheet selected."
        cols = list(conv.sheets_data[sheet].columns)[:MAX_COLUMNS]
        return f"Settings applied for **{sheet}** ({len(cols)} columns)."

    # Collect all dropdown inputs for apply
    _apply_inputs = [types_tab_sheet]
    for entry in col_config_rows:
        _apply_inputs.extend([entry["hier"], entry["delim"], entry["group"], entry["type"]])

    adv_apply_btn.click(
        _apply_col_settings,
        inputs=_apply_inputs,
        outputs=[adv_status],
    )

    # ── Sheet tab switch — auto-save previous sheet, then load new one ──

    def on_types_tab_change(new_sheet, prev_sheet, *dropdown_values):
        # Auto-save the PREVIOUS sheet's dropdown values before switching
        _save_dropdowns_to_config(prev_sheet, dropdown_values)

        if not new_sheet or new_sheet not in conv.sheets_data:
            lvl_update = gr.update(value="0")
            return [lvl_update, new_sheet] + _refresh_col_config_outputs(None)
        lvl = conv.adv_sheet_levels.get(new_sheet, 0)
        return [gr.update(value=str(lvl)), new_sheet] + _refresh_col_config_outputs(new_sheet)

    # Inputs: new sheet selection, previous sheet state, then all dropdowns
    _sheet_change_inputs = [types_tab_sheet, adv_sheet_state]
    for entry in col_config_rows:
        _sheet_change_inputs.extend([entry["hier"], entry["delim"], entry["group"], entry["type"]])

    types_tab_sheet.change(
        on_types_tab_change,
        inputs=_sheet_change_inputs,
        outputs=[hier_level_count, adv_sheet_state] + _col_cfg_outputs,
    )

    # ── Hierarchy level count change — update dropdown choices ──

    def on_hier_count_change(val, sheet):
        if not sheet or sheet not in conv.sheets_data:
            return _refresh_col_config_outputs(None)
        lvl = int(val) if val else 0
        conv.adv_sheet_levels[sheet] = lvl
        # Reset any columns that exceed the new max level
        for col, cfg in conv.adv_col_config.get(sheet, {}).items():
            if cfg.get("hier", 0) > lvl:
                cfg["hier"] = 0
        return _refresh_col_config_outputs(sheet)

    hier_level_count.change(
        on_hier_count_change,
        inputs=[hier_level_count, types_tab_sheet],
        outputs=_col_cfg_outputs,
    )

    # ── Grouping toggle ──

    def toggle_adv_grouping(is_enabled, sheet):
        conv.adv_grouping_enabled = is_enabled
        return [gr.update(visible=is_enabled)] + _refresh_col_config_outputs(sheet)

    adv_group_enable_cb.change(
        toggle_adv_grouping,
        inputs=[adv_group_enable_cb, types_tab_sheet],
        outputs=[adv_group_manager] + _col_cfg_outputs,
    )

    def add_adv_group(new_name, sheet):
        name = str(new_name).strip()
        if name and name not in conv.adv_groups:
            conv.adv_groups.append(name)
        return ([gr.update(value=""), gr.update(choices=conv.adv_groups)]
                + _refresh_col_config_outputs(sheet))

    adv_add_group_btn.click(
        add_adv_group,
        inputs=[adv_new_group_name, types_tab_sheet],
        outputs=[adv_new_group_name, adv_del_group_dropdown] + _col_cfg_outputs,
    )

    def remove_adv_group(del_name, sheet):
        name = str(del_name).strip()
        if name in conv.adv_groups:
            conv.adv_groups.remove(name)
            for s_cfg in conv.adv_col_config.values():
                for cfg in s_cfg.values():
                    if cfg.get("group") == name:
                        cfg["group"] = "None"
        return ([gr.update(choices=conv.adv_groups, value=None)]
                + _refresh_col_config_outputs(sheet))

    adv_del_group_btn.click(
        remove_adv_group,
        inputs=[adv_del_group_dropdown, types_tab_sheet],
        outputs=[adv_del_group_dropdown] + _col_cfg_outputs,
    )

    # ── Mode toggle ──

    def toggle_mode(mode):
        is_adv = (mode == "Advanced")
        return (
            gr.update(visible=is_adv),      # advanced_step1_col
            gr.update(visible=not is_adv),   # simple_hier_group
            gr.update(visible=is_adv),       # advanced_group
        )

    mode_toggle.change(toggle_mode, inputs=[mode_toggle],
                       outputs=[advanced_step1_col, simple_hier_group, advanced_group])

    # ── Load file / GSheet ──

    def update_ui_after_load(ok, msg, sheets):
        if not ok:
            empty_col_cfg = _refresh_col_config_outputs(None)
            return ([
                msg,
                gr.update(choices=[]),        # sheets_sel
                gr.update(choices=[]),        # hier_sheet
                gr.update(choices=[]),        # nt_sheet
                gr.update(choices=[]),        # types_tab_sheet
                gr.update(value="0"),         # hier_level_count
                [],                           # nt_view
                invalidate(),                 # gen_file
                gr.update(value=True),        # use_file_root_cb
                gr.update(choices=[]),        # nt_col
                [],                           # hier_data
                "",                           # adv_sheet_state
            ] + empty_col_cfg)

        conv.selected_sheets = set(s for s in sheets if s != "structure")
        sel_list = sorted(list(conv.selected_sheets))
        first = sel_list[0] if sel_list else None
        new_root_state = (len(sel_list) > 1)
        conv.use_file_root = new_root_state

        for s in sel_list:
            if s in conv.sheets_data and not conv.sheet_title_cols.get(s):
                cols = list(conv.sheets_data[s].columns)
                if cols:
                    conv.sheet_title_cols[s] = cols[0]

        return ([
            msg,
            gr.update(choices=sheets, value=list(conv.selected_sheets)),
            gr.update(choices=sel_list, value=first),
            gr.update(choices=sel_list, value=first),
            gr.update(choices=sel_list, value=first),
            gr.update(value=str(conv.adv_sheet_levels.get(first, 0)) if first else "0"),
            format_nt_df(),
            gr.update(value=None, visible=False),
            gr.update(value=new_root_state),
            gr.update(choices=list(conv.sheets_data[first].columns) if first else []),
            [],
            first or "",
        ] + _refresh_col_config_outputs(first))

    def do_load_file(file):
        if not file:
            return update_ui_after_load(False, "No file", [])
        ok, msg, ch = conv.load_xlsx(file.name)
        return update_ui_after_load(ok, msg, ch)

    def do_load_gs(url, json_txt):
        ok, msg, ch = conv.load_gsheet(url, json_txt)
        return update_ui_after_load(ok, msg, ch)

    common_outputs = [
        status, sheets_sel, hier_sheet, nt_sheet, types_tab_sheet,
        hier_level_count, nt_view, gen_file, use_file_root_cb, nt_col,
        hier_data, adv_sheet_state,
    ] + _col_cfg_outputs

    load_file_btn.click(do_load_file, inputs=[up_file], outputs=common_outputs)
    load_gs_btn.click(do_load_gs, inputs=[gs_url, gs_json], outputs=common_outputs)

    # ── Sheet selection change ──

    def on_save_sel(sel):
        conv.selected_sheets = set(sel or [])
        new_choices = sorted(list(conv.selected_sheets))
        conv.hierarchy_configs = {k: v for k, v in conv.hierarchy_configs.items()
                                  if k in new_choices}
        conv.sheet_title_cols = {k: v for k, v in conv.sheet_title_cols.items()
                                 if k in new_choices}
        first = new_choices[0] if new_choices else None
        return ([
            gr.update(choices=new_choices, value=first),
            gr.update(choices=new_choices, value=first),
            gr.update(choices=new_choices, value=first),
            gr.update(value=str(conv.adv_sheet_levels.get(first, 0)) if first else "0"),
            format_nt_df(),
            invalidate(),
        ] + _refresh_col_config_outputs(first))

    sheets_sel.change(
        on_save_sel,
        inputs=[sheets_sel],
        outputs=[hier_sheet, nt_sheet, types_tab_sheet, hier_level_count,
                 nt_view, gen_file] + _col_cfg_outputs,
    )

    # ── Simple Mode Hierarchy ──

    def on_hier_sheet_change(sheet):
        if not sheet or sheet not in conv.sheets_data:
            return gr.update(choices=[]), []
        cols = list(conv.sheets_data[sheet].columns)
        current = conv.hierarchy_configs.get(sheet, [])
        used_cols = [x['col'] if isinstance(x, dict) else str(x) for x in current]
        avail = [c for c in cols if c not in used_cols]
        return gr.update(choices=avail), format_hier_df(sheet)

    def add_hier_level(sheet, col, delim):
        if not sheet or not col:
            return format_hier_df(sheet), gr.update(), invalidate()
        current = conv.hierarchy_configs.get(sheet, [])
        d = None if delim == "No Separation" else delim
        current.append({'col': col, 'delim': d})
        conv.hierarchy_configs[sheet] = current
        cols = list(conv.sheets_data[sheet].columns)
        used_cols = [x['col'] if isinstance(x, dict) else str(x) for x in current]
        avail = [c for c in cols if c not in used_cols]
        return format_hier_df(sheet), gr.update(choices=avail, value=None), invalidate()

    hier_sel = gr.State(-1)

    def on_hier_select(evt: gr.SelectData):
        return evt.index[0]

    def hier_delete(sheet, idx):
        if not sheet or idx < 0:
            return format_hier_df(sheet), gr.update(), invalidate()
        current = conv.hierarchy_configs.get(sheet, [])
        if idx < len(current):
            current.pop(idx)
            conv.hierarchy_configs[sheet] = current
        cols = list(conv.sheets_data[sheet].columns)
        used_cols = [x['col'] if isinstance(x, dict) else str(x) for x in current]
        avail = [c for c in cols if c not in used_cols]
        return format_hier_df(sheet), gr.update(choices=avail), invalidate()

    def hier_move(sheet, idx, direction):
        if not sheet:
            return format_hier_df(sheet), gr.update(), invalidate()
        cur = conv.hierarchy_configs.get(sheet, [])
        if not cur:
            return format_hier_df(sheet), gr.update(), invalidate()
        if direction == "up" and idx > 0:
            cur[idx], cur[idx - 1] = cur[idx - 1], cur[idx]
        elif direction == "down" and idx < len(cur) - 1:
            cur[idx], cur[idx + 1] = cur[idx + 1], cur[idx]
        conv.hierarchy_configs[sheet] = cur
        return format_hier_df(sheet), gr.update(), invalidate()

    hier_sheet.change(on_hier_sheet_change, inputs=[hier_sheet],
                      outputs=[hier_add_col, hier_data])
    hier_add_btn.click(add_hier_level, inputs=[hier_sheet, hier_add_col, hier_delim],
                       outputs=[hier_data, hier_add_col, gen_file])
    hier_data.select(on_hier_select, None, hier_sel)
    hier_up_btn.click(lambda s, i: hier_move(s, i, "up"),
                      inputs=[hier_sheet, hier_sel], outputs=[hier_data, gen_file])
    hier_down_btn.click(lambda s, i: hier_move(s, i, "down"),
                        inputs=[hier_sheet, hier_sel], outputs=[hier_data, gen_file])
    hier_del_btn.click(hier_delete, inputs=[hier_sheet, hier_sel],
                       outputs=[hier_data, hier_add_col, gen_file])

    # ── Misc bindings ──

    def set_root_flag(val):
        conv.use_file_root = val
        return invalidate()

    def set_direction(val):
        conv.edge_direction = val
        return invalidate()

    use_file_root_cb.change(set_root_flag, inputs=[use_file_root_cb], outputs=[gen_file])
    direction.change(set_direction, inputs=[direction], outputs=[gen_file])
    hier_show_class_l1.change(invalidate, outputs=[gen_file])
    hier_show_class_deep.change(invalidate, outputs=[gen_file])
    global_data_delim.change(invalidate, outputs=[gen_file])

    def on_nt_sheet_change(sheet):
        if not sheet or sheet not in conv.sheets_data:
            return gr.update(choices=[])
        return gr.update(choices=list(conv.sheets_data[sheet].columns))

    def set_node_title(sheet, col):
        if sheet and col:
            conv.sheet_title_cols[sheet] = col
        return format_nt_df(), invalidate()

    nt_sheet.change(on_nt_sheet_change, inputs=[nt_sheet], outputs=[nt_col])
    nt_btn.click(set_node_title, inputs=[nt_sheet, nt_col], outputs=[nt_view, gen_file])

    # ── Generate ──

    def gen_xml(mode_val, use_l1, use_deep, data_delim, current_sheet, *dropdown_values):
        # Auto-save the currently visible sheet's dropdowns before generating
        _save_dropdowns_to_config(current_sheet, dropdown_values)

        conv.use_prefix_l1 = use_l1
        conv.use_prefix_deep = use_deep
        conv.data_delimiter = None if data_delim == "No Separation" else data_delim
        try:
            path, n, m = conv.generate(mode_val)
            return gr.update(value=path, visible=True), f"Success! {n} nodes, {m} edges generated."
        except Exception as e:
            traceback.print_exc()
            return gr.update(visible=False), f"**Error:** {str(e)}"

    _gen_inputs = [mode_toggle, hier_show_class_l1, hier_show_class_deep,
                   global_data_delim, adv_sheet_state]
    for entry in col_config_rows:
        _gen_inputs.extend([entry["hier"], entry["delim"], entry["group"], entry["type"]])

    gen_btn.click(
        gen_xml,
        inputs=_gen_inputs,
        outputs=[gen_file, gen_status],
    )

    # ── Persist Google Sheets creds in browser localStorage ──
    app.load(None, None, gs_json, js="() => localStorage.getItem('gs_creds') || ''")
    gs_json.change(None, gs_json, None, js="(v) => localStorage.setItem('gs_creds', v)")

if __name__ == "__main__":
    app.launch()
