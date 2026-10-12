#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Data App -> EDA automation.  SINGLE FILE: engine + customtkinter UI.

    python "dataapp_eda_automation.py"

How it works
------------
1. You press "1. Open browser".  Selenium launches Chrome and navigates to the
   Data App.  That Chrome window is the ONLY browser used from then on.
2. In that window YOU do the manual part: SSO login, API token, Submit,
   Organization, schema, and click "Load Data".
3. You come back to this window and press "2. Platform loaded - Run".
   Everything after Load Data is automated:

       Objective tab -> Date Range -> Plots -> Time-Series Plots
       -> Marker type -> [x] Individual Plots -> [x] Impute NA to 0 -> Go Plot
       -> per chart: hide 'normalizationid',
                     show only 'count' -> download PNG,
                     show only 'value' -> download PNG
       -> read the real series out of the Plotly figure
       -> write a Word doc: heading (variable name), both images,
          and observations (peaks, dips, zero runs, trend, anomalies)

Everything the automation touches is looked up through the SELECTORS table
below, which holds ordered fallback locators.  Tick "Dump DOM" to write a JSON
inventory of every input/button/tab so the locators can be pinned exactly.

Requires: selenium, python-docx, customtkinter (all already installed).
"""

from __future__ import annotations

import json
import os
import queue
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk
from selenium import webdriver
from selenium.common.exceptions import (
    ElementClickInterceptedException,
    JavascriptException,
    NoSuchWindowException,
    StaleElementReferenceException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.chrome.options import Options as ChromeOptions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select

# --------------------------------------------------------------------------- #
# Constants
# --------------------------------------------------------------------------- #

DEFAULT_URL = "https://dashboards.production.bw7.io/data/"
DEFAULT_PROFILE = Path.home() / ".dataapp_chrome_profile"
SETTINGS_PATH = Path.home() / ".dataapp_gui_settings.json"

X = By.XPATH
C = By.CSS_SELECTOR

# Ordered candidate locators -- the first one that resolves wins.
# Once the real DOM ids are known, prepend them to the relevant list.
SELECTORS: dict[str, list[tuple[str, str]]] = {
    # ---- top-level tabs -------------------------------------------------- #
    "tab_summary": [
        (X, "//a[@data-toggle='tab' or @data-bs-toggle='tab'][normalize-space()='Summary']"),
        (X, "//*[normalize-space()='Summary'][self::a or self::button]"),
    ],
    "tab_objective": [
        (X, "//a[@data-toggle='tab' or @data-bs-toggle='tab'][normalize-space()='Objective']"),
        (X, "//ul[contains(@class,'nav')]//a[normalize-space()='Objective']"),
        (X, "//*[normalize-space()='Objective'][self::a or self::button]"),
    ],
    "tab_nonmedia": [
        (X, "//a[@data-toggle='tab' or @data-bs-toggle='tab']"
            "[normalize-space()='Non-media' or normalize-space()='Non media']"),
        (X, "//*[contains(normalize-space(),'Non-media')][self::a or self::button]"),
    ],
    "tab_media": [
        (X, "//a[@data-toggle='tab' or @data-bs-toggle='tab'][normalize-space()='Media']"),
        (X, "//*[normalize-space()='Media'][self::a or self::button]"),
    ],
    # ---- date range ------------------------------------------------------ #
    # Only a fallback: JS_DATE_INPUTS below is what normally resolves these,
    # because the two ends live in one wrapper whose markup differs between
    # Bootstrap 3 and 5, and positional XPath quietly picks the wrong one.
    "date_from": [
        (C, ".shiny-date-range-input input[type=text]:first-of-type"),
        (C, ".input-daterange input:first-of-type"),
        (X, "//*[contains(@class,'shiny-date-range-input')]//input[1]"),
        (X, "//label[contains(.,'Date Range')]/following::input[1]"),
        (C, "input.datepicker"),
    ],
    "date_to": [
        (C, ".shiny-date-range-input input[type=text]:last-of-type"),
        (C, ".input-daterange input:last-of-type"),
        (X, "//*[contains(@class,'shiny-date-range-input')]//input[2]"),
        (X, "//label[contains(.,'Date Range')]/following::input[2]"),
    ],
    # ---- plot sub-tabs --------------------------------------------------- #
    "tab_plots": [
        (X, "//a[@data-toggle='tab' or @data-bs-toggle='tab'][normalize-space()='Plots']"),
        (X, "//*[normalize-space()='Plots'][self::a or self::button]"),
    ],
    "tab_timeseries": [
        (X, "//a[contains(normalize-space(),'Time-Series Plots') or "
            "contains(normalize-space(),'Time Series Plots')]"),
        (X, "//*[contains(normalize-space(),'Time-Series')][self::a or self::button]"),
    ],
    # ---- plot options ---------------------------------------------------- #
    "go_plot": [
        (C, "button#go_plot, button#goplot, button#goPlot, #plotBtn"),
        (X, "//button[contains(.,'Go Plot')]"),
        (X, "//*[self::a or self::div][contains(@class,'btn')][contains(.,'Go Plot')]"),
    ],
    # ---- confirmation that Load Data really ran -------------------------- #
    "loaded_marker": [
        (X, "//*[normalize-space()='Objective'][self::a or self::button]"),
    ],
}

# Sidebar controls. Only used to rebuild the session after a Shiny disconnect --
# the normal flow has the user fill these in by hand.
SELECTORS.update({
    "api_token": [
        (C, "input#token, input#api_token, input#apiToken"),
        (X, "//label[contains(translate(.,'API TOKEN','api token'),'api token')]"
            "/following::input[@type='password' or @type='text'][1]"),
        (C, "input[type='password']"),
    ],
    "submit_token": [
        (C, "button#submit, button#submit_token"),
        (X, "//button[normalize-space()='Submit' or contains(.,'Submit')]"),
    ],
    "load_data": [
        (C, "button#load, button#load_data, button#loadData"),
        (X, "//button[contains(.,'Load Data')]"),
        (X, "//*[self::a or self::div][contains(@class,'btn')][contains(.,'Load Data')]"),
    ],
})

# Where the report lands: Downloads / data sanity / <project> /
DOWNLOADS_ROOT = Path.home() / "Downloads" / "data sanity"


def radio_locators(label: str) -> list[tuple[str, str]]:
    """Marker-type radios are matched by their visible label text."""
    return [
        (X, f"//label[normalize-space()='{label}']/input[@type='radio']"),
        (X, f"//input[@type='radio'][@value='{label}' or @value='{label.lower()}']"),
        (X, f"//*[normalize-space()='{label}']/preceding-sibling::input[@type='radio'][1]"),
        (X, f"//*[normalize-space()='{label}']/input[@type='radio']"),
        (X, f"//label[normalize-space()='{label}']"),
    ]


def norm(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", (s or "").lower())


# --------------------------------------------------------------------------- #
# Per-tab behaviour lives HERE, not in the UI. Each tab of the Data App needs
# different plot options, different traces and a different way of grouping the
# charts into report boxes, and none of that is a user decision.
# --------------------------------------------------------------------------- #

@dataclass(frozen=True)
class TabProfile:
    heading: str                      # Heading 1 text in the Word report
    marker: str = "Lines"
    # checkbox label -> wanted state. Each tab has its own set: Objective and
    # Non-media have Individual Plots / Impute NA to 0, Media has the dubious
    # data switches instead.
    checkboxes: tuple[tuple[str, bool], ...] = ()
    iterate_feeds: bool = False       # walk every option of "Feed Name" in turn
    feed_label: str = "feed name"
    filter_values: tuple[str, ...] = ()   # pick these in "Filter Values"
    export: tuple[str, ...] | None = None  # trace order; None = all on the chart
    hide: tuple[str, ...] = ("normalizationid",)
    group_mode: str = "by_variable"   # "single" | "by_variable" | "by_channel"
    data_type: str = ""               # box label when group_mode == "single"
    settle: float = 1.0               # seconds the chart count must hold still
    plot_timeout: int = 240


TAB_PROFILES: dict[str, TabProfile] = {
    "Objective": TabProfile(
        heading="OBJECTIVE",
        checkboxes=(("Individual Plots", True), ("Impute NA to 0", True)),
        export=("count", "value"),
        hide=("normalizationid",),
        group_mode="single",
        data_type="OBJECTIVE",
    ),
    "Non-media": TabProfile(
        heading="NON-MEDIA DATA",
        checkboxes=(("Individual Plots", True), ("Impute NA to 0", True)),
        iterate_feeds=True,
        export=None,
        hide=("normalizationid",),
        group_mode="by_feed",
    ),
    # Media has no Individual Plots / Impute NA. It filters on values instead,
    # renders many charts at once, and each chart carries a 'net' and an
    # 'impressions' series -- exported impressions first, then net.
    "Media": TabProfile(
        heading="MEDIA DATA",
        checkboxes=(("Remove dubious data", True),),
        filter_values=("net + impressions",),
        export=("impressions", "net"),
        hide=(),
        group_mode="by_channel",
        settle=5.0,
        plot_timeout=900,
    ),
}

SECTION_ORDER = ("Objective", "Non-media", "Media")


def checkbox_locators(label: str) -> list[tuple[str, str]]:
    return [
        (X, f"//label[contains(normalize-space(),'{label}')]//input[@type='checkbox']"),
        (X, f"//*[contains(normalize-space(text()),'{label}')]"
            f"/preceding::input[@type='checkbox'][1]"),
        (X, f"//span[contains(normalize-space(),'{label}')]"
            f"/preceding-sibling::input[@type='checkbox']"),
    ]


def channel_of(title: str) -> str:
    """
    Media charts are titled from the UID, e.g. 'flavors_googleads_paidsearch_tra'
    or 'flavors_meta_paidsocial_awa'. The channel is the token beginning with
    'paid'; if there is none, fall back to the 4th token, then the 3rd.
    """
    parts = [p for p in title.split("_") if p]
    for p in parts:
        if p.lower().startswith("paid"):
            return p
    if len(parts) >= 4:
        return parts[3]
    if len(parts) >= 3:
        return parts[2]
    return title


def split_title(title: str) -> tuple[str, str]:
    """
    Non-media charts are titled '<variable>_<uid>', e.g.
    'dispwofeatincrcaseper_bulleitwhiskey'. Split on the LAST underscore, so
    variables that themselves contain underscores still group correctly.
    """
    if "_" in title:
        var, uid = title.rsplit("_", 1)
        return var, uid
    return title, ""


# The app is fronted by oauth2-proxy, which bounces to Keycloak
# (keycloak.production2.bw7.io/auth/realms/Internal_Applications/...).
# These fragments identify "you are still on a login page, not the app".
LOGIN_URL_HINTS = (
    "/oauth2/", "sign_in", "signin", "openid-connect", "keycloak", "/realms/",
    "login.microsoftonline", "okta.com", "accounts.google.com",
)


def is_login_page(url: str) -> bool:
    u = (url or "").lower()
    return any(h in u for h in LOGIN_URL_HINTS)


# --------------------------------------------------------------------------- #
# Config
# --------------------------------------------------------------------------- #

@dataclass
class Config:
    url: str = DEFAULT_URL
    # Blank on purpose: blank means "use the range the Data App is showing".
    # The app derives that from the data it just loaded, so it is the only
    # range that can be right for every model; the UI boxes are an override,
    # not the source.
    date_from: str = ""
    date_to: str = ""
    # every range actually seen on the page: (from, to, tag), for the report
    dates_seen: list = field(default_factory=list)
    sections: tuple[str, ...] = ("Objective",)   # which Data App tabs to capture
    project: str = ""                     # names the Downloads sub-folder
    brand: str = ""
    outdir: Path | None = None            # None -> Downloads/data sanity/<project>
    run_dir: Path | None = None           # the per-run folder, set at run time
    template: Path | None = None
    delivered_by: str = ""
    profile_dir: Path | None = field(default_factory=lambda: DEFAULT_PROFILE)
    img_width: int = 1400
    img_height: int = 700
    timeout: int = 30            # how long to hunt for one element
    idle_timeout: int = 20       # cap on "wait for the page to settle"
    plot_timeout: int = 240      # how long Go Plot may take to render
    debug_dom: bool = False
    # runtime hooks (set by the UI)
    org: str = ""
    schema: str = ""
    token_cache: str = ""            # read off the page, used only to reconnect
    org_value: str = ""              # option values, so a rebuild can re-pick
    schema_value: str = ""
    reconnect_wait: int = 900        # how long to wait for a manual rescue
    on_log: object = None
    on_progress: object = None       # callable(dict) for the UI progress bar
    cancel: object = None


# --------------------------------------------------------------------------- #
# Plotly JS
# --------------------------------------------------------------------------- #

# Shiny keeps every tab's content in the DOM and merely hides the inactive
# panes, so a plain querySelectorAll('.js-plotly-plot') also returns charts from
# tabs we are not on. That made Non-media re-export the Objective chart. Every
# script below therefore works through visiblePlots(), never the raw selector.
JS_PLOTS = """
function visiblePlots() {
  return Array.from(document.querySelectorAll('.js-plotly-plot')).filter(
    function (g) {
      if (g.offsetParent === null &&
          getComputedStyle(g).position !== 'fixed') return false;
      var r = g.getBoundingClientRect();
      if (!r.width || !r.height) return false;
      var pane = g.closest('.tab-pane');
      if (pane && !pane.classList.contains('active')) return false;
      return true;
    });
}
"""

JS_TRACE_INFO = JS_PLOTS + """
const out = [];
visiblePlots().forEach((g, gi) => {
  const t = (g.layout && g.layout.title && (g.layout.title.text || g.layout.title)) || '';
  out.push({
    index: gi,
    title: String(t),
    traces: (g.data || []).map((tr, ti) => ({
      i: ti,
      name: tr.name || ('trace ' + ti),
      visible: (tr.visible === undefined ? true : tr.visible),
      npoints: (tr.y || tr.x || []).length
    }))
  });
});
return out;
"""

JS_SET_VISIBILITY = JS_PLOTS + """
const gi = arguments[0];
const keep = arguments[1].map(s => s.toLowerCase().replace(/[^a-z0-9]/g, ''));
const g = visiblePlots()[gi];
if (!g) return {ok: false, reason: 'no-plot'};
const key = t => (t.name || '').toLowerCase().replace(/[^a-z0-9]/g, '');
const matched = (g.data || []).filter(t => keep.includes(key(t))).length;
// Never hide every series: a blank chart exports a blank PNG, and that always
// means something is wrong (stale index, renamed trace) rather than intended.
if (!matched) {
  return {ok: false, reason: 'no-match',
          have: (g.data || []).map(t => t.name || '')};
}
const vis = (g.data || []).map(t => keep.includes(key(t)) ? true : 'legendonly');
Plotly.restyle(g, {visible: vis});
const title = (g.layout && g.layout.title &&
              (g.layout.title.text || g.layout.title)) || '';
return {ok: true, matched: matched, total: (g.data || []).length,
        title: String(title)};
"""

# 100+ charts can be re-rendered while we work through them, which shuffles
# indices. Re-resolve a chart by its title before touching it.
JS_FIND_BY_TITLE = JS_PLOTS + """
const want = String(arguments[0]);
const gs = visiblePlots();
for (let i = 0; i < gs.length; i++) {
  const g = gs[i];
  const t = (g.layout && g.layout.title &&
            (g.layout.title.text || g.layout.title)) || '';
  if (String(t) === want) return i;
}
return -1;
"""

JS_EXTRACT_DATA = JS_PLOTS + """
const gi = arguments[0];
const g = visiblePlots()[gi];
if (!g) return null;
const t = (g.layout && g.layout.title && (g.layout.title.text || g.layout.title)) || '';
return {
  title: String(t),
  traces: (g.data || []).map(tr => ({
    name: tr.name || '',
    x: (tr.x || []).map(v => (v instanceof Date) ? v.toISOString() : String(v)),
    y: (tr.y || []).map(v => (v === null || v === undefined) ? null : Number(v))
  }))
};
"""

JS_DOWNLOAD_IMAGE = JS_PLOTS + """
const done = arguments[arguments.length - 1];
const gi = arguments[0], fname = arguments[1], w = arguments[2], h = arguments[3];
const g = visiblePlots()[gi];
if (!g) { done('no-plot'); return; }
Plotly.downloadImage(g, {format: 'png', filename: fname, width: w, height: h})
      .then(() => done('ok')).catch(e => done('err:' + e));
"""

# The Non-media "Feed Name" box is a SINGLE select: one feed at a time. So we
# read every option first, then work through them one by one.
JS_LIST_OPTIONS = """
const wanted = arguments[0].toLowerCase();
function labelOf(el) {
  const grp = el.closest('.form-group, .shiny-input-container');
  const lab = grp && grp.querySelector('label, .control-label');
  return lab ? lab.innerText.trim().toLowerCase() : '';
}
for (const sel of document.querySelectorAll('select')) {
  if (!labelOf(sel).includes(wanted)) continue;
  let opts;
  if (sel.selectize) {
    const o = sel.selectize.options;
    opts = Object.keys(o).map(v => ({value: v, text: o[v].text || v}));
  } else {
    opts = Array.from(sel.options).map(o => ({value: o.value, text: o.text}));
  }
  opts = opts.filter(o => o.value !== '' && o.value !== null);
  const cur = sel.selectize ? sel.selectize.getValue()
                            : (sel.value || '');
  return {found: true, label: labelOf(sel), multiple: !!sel.multiple,
          widget: sel.selectize ? 'selectize' : 'select',
          options: opts, current: cur};
}
return {found: false, options: [], current: ''};
"""

JS_SELECT_ONE = """
const value = String(arguments[0]);
const wanted = arguments[1].toLowerCase();
function labelOf(el) {
  const grp = el.closest('.form-group, .shiny-input-container');
  const lab = grp && grp.querySelector('label, .control-label');
  return lab ? lab.innerText.trim().toLowerCase() : '';
}
for (const sel of document.querySelectorAll('select')) {
  if (!labelOf(sel).includes(wanted)) continue;
  if (sel.selectize) {
    sel.selectize.setValue(value);            // fires change
    return {ok: true, now: sel.selectize.getValue()};
  }
  sel.value = value;
  sel.dispatchEvent(new Event('change', {bubbles: true}));
  if (window.jQuery) jQuery(sel).trigger('change');
  return {ok: true, now: sel.value};
}
return {ok: false, now: ''};
"""

# Pick specific options in a labelled selector (the Media "Filter Values" box).
JS_SELECT_VALUES = """
const wanted = arguments[0].map(w => w.toLowerCase().replace(/[^a-z0-9]/g, ''));
const lblWant = arguments[1].toLowerCase();

function labelOf(el) {
  const grp = el.closest('.form-group, .shiny-input-container');
  const lab = grp && grp.querySelector('label, .control-label');
  return lab ? lab.innerText.trim().toLowerCase() : '';
}
function key(s) { return String(s).toLowerCase().replace(/[^a-z0-9]/g, ''); }

for (const sel of document.querySelectorAll('select')) {
  if (!labelOf(sel).includes(lblWant)) continue;
  const before = Array.from(sel.selectedOptions || []).map(o => o.value);

  if (sel.selectize) {
    const opts = sel.selectize.options;
    const hits = Object.keys(opts).filter(v =>
        wanted.includes(key(v)) || wanted.includes(key(opts[v].text || '')));
    if (hits.length) sel.selectize.setValue(hits);
    return {mode: 'selectize', label: labelOf(sel),
            available: Object.keys(opts).map(v => opts[v].text || v),
            before: before, after: sel.selectize.getValue(), matched: hits.length};
  }
  const all = Array.from(sel.options);
  const hits = all.filter(o => wanted.includes(key(o.value)) ||
                               wanted.includes(key(o.text)));
  if (hits.length) {
    for (const o of all) o.selected = false;
    for (const o of hits) o.selected = true;
    sel.dispatchEvent(new Event('change', {bubbles: true}));
    if (window.jQuery) jQuery(sel).trigger('change');
  }
  return {mode: sel.multiple ? 'multiple' : 'single-select', label: labelOf(sel),
          available: all.map(o => o.text), before: before,
          after: hits.map(o => o.value), matched: hits.length};
}
return {mode: 'not-found', label: '', available: [], before: [], after: [],
        matched: 0};
"""

# Read the org / schema the user picked, so the report names itself.
JS_READ_CONTEXT = """
function valueNear(words) {
  for (const sel of document.querySelectorAll('select')) {
    const grp = sel.closest('.form-group, .shiny-input-container');
    const lab = grp && grp.querySelector('label, .control-label');
    const txt = (lab ? lab.innerText : '').trim().toLowerCase();
    if (words.some(w => txt === w || txt.startsWith(w))) {
      if (sel.selectize) return String(sel.selectize.getValue() || '');
      const o = sel.selectedOptions && sel.selectedOptions[0];
      return o ? (o.text || o.value) : (sel.value || '');
    }
  }
  return '';
}
return {
  organization: valueNear(['organization', 'organisation', 'org']),
  schema: valueNear(['schema']),
  feed: valueNear(['feed name', 'feed'])
};
"""

# The Date Range box is the app's own statement of what it loaded, so it is
# read, never assumed. Finding it positionally is what used to fail: both ends
# sit in one wrapper, every tab keeps its own hidden copy, and Bootstrap 3 and 5
# nest the inputs differently. This walks the wrappers instead and prefers the
# one on the tab that is actually showing.
JS_DATE_COMMON = r"""
function _paneActive(el) {
  let node = el;
  while (node && node.closest) {
    const pane = node.closest('.tab-pane');
    if (!pane) break;
    if (!pane.classList.contains('active')) return false;
    node = pane.parentElement;
  }
  return true;
}
function _shown(el) {
  if (!el) return false;
  if (el.offsetParent === null &&
      getComputedStyle(el).position !== 'fixed') return false;
  const r = el.getBoundingClientRect();
  return !!(r.width && r.height);
}
function _dateBoxes() {
  const out = [];
  for (const b of document.querySelectorAll(
        '.shiny-date-range-input, .input-daterange')) {
    if (out.indexOf(b) < 0) out.push(b);
  }
  // apps that style their own picker: go by the label instead
  for (const lab of document.querySelectorAll('label, .control-label')) {
    if (!/date\s*range/i.test(lab.textContent || '')) continue;
    const grp = lab.closest('.form-group, .shiny-input-container') ||
                lab.parentElement;
    if (grp && out.indexOf(grp) < 0) out.push(grp);
  }
  return out;
}
function _dateInputs(box) {
  return Array.from(box.querySelectorAll('input')).filter(function (i) {
    const t = (i.type || 'text').toLowerCase();
    return t !== 'hidden' && t !== 'checkbox' && t !== 'radio' &&
           t !== 'password' && t !== 'button';
  });
}
function _bestDateBox() {
  let best = null, bestScore = -1;
  for (const b of _dateBoxes()) {
    const ins = _dateInputs(b);
    if (ins.length < 2) continue;
    const score = (_paneActive(b) ? 2 : 0) + (_shown(b) ? 1 : 0);
    if (score > bestScore) {
      best = {box: b, ins: ins, score: score};
      bestScore = score;
    }
  }
  return best;
}
"""

# Values only -- safe to call at any time, including before a tab is opened.
JS_DATE_RANGE = JS_DATE_COMMON + """
const best = _bestDateBox();
if (!best) return null;
function attr(el, names) {
  for (const n of names) {
    const v = el.getAttribute(n);
    if (v) return v;
  }
  return '';
}
const a = best.ins[0], z = best.ins[1];
return {
  from: (a.value || '').trim(),
  to:   (z.value || '').trim(),
  min:  attr(a, ['data-min-date', 'data-date-start-date', 'min']),
  max:  attr(z, ['data-max-date', 'data-date-end-date', 'max']),
  fmt:  attr(a, ['data-date-format', 'data-format']),
  label: (function () {
           const l = best.box.querySelector('label, .control-label');
           return l ? (l.textContent || '').trim() : '';
         })(),
  visible: best.score >= 1,
  active:  best.score >= 2
};
"""

# The two <input> elements themselves, so typing goes to the right end.
JS_DATE_INPUTS = JS_DATE_COMMON + """
const best = _bestDateBox();
return best ? [best.ins[0], best.ins[1]] : null;
"""

JS_GET_PLOT = JS_PLOTS + """
return visiblePlots()[arguments[0]] || null;
"""

JS_COUNT_PLOTS = JS_PLOTS + """
const gs = visiblePlots();
if (!gs.length) return {n: 0, ready: false};
for (const g of gs) { if (!g.data || !g.data.length) return {n: gs.length, ready: false}; }
return {n: gs.length, ready: true};
"""

# Every tab carries its OWN Marker type radios, checkboxes, Go Plot button and
# date inputs, and inactive panes stay in the DOM. An unscoped locator therefore
# returns the Objective tab's hidden controls and the visible tab never changes.
# This rejects anything inside a pane that is not active, checking every
# ancestor pane because the Plots sub-tabs nest inside the main ones.
# Shiny drops its websocket when the server recycles the session; the page then
# greys out behind a "Disconnected from the server / Reload" overlay and every
# control stops responding. Nothing we do causes it, but we must survive it.
JS_DISCONNECTED = """
function shown(el) {
  if (!el) return false;
  const r = el.getBoundingClientRect();
  if (!r.width || !r.height) return false;
  const st = getComputedStyle(el);
  if (st.display === 'none' || st.visibility === 'hidden') return false;
  return parseFloat(st.opacity || '1') > 0.05;
}
for (const id of ['shiny-disconnected-overlay', 'ss-connect-dialog',
                  'shiny-disconnected']) {
  if (shown(document.getElementById(id))) return 'id:' + id;
}
if (shown(document.querySelector('.shiny-disconnected'))) return 'class';
if (document.body && document.body.classList.contains('disconnected'))
  return 'body-class';
// Last resort, and the one that actually catches Shiny Server: any visible
// element whose own text says it. Checked on small leaf nodes only, so the
// whole <body> never matches.
const rx = /disconnected from the server/i;
for (const el of document.querySelectorAll('div, p, span, section')) {
  if (el.children.length > 3) continue;
  const t = (el.textContent || '');
  if (t.length < 200 && rx.test(t) && shown(el)) return 'text';
}
try {
  const sk = window.Shiny && Shiny.shinyapp && Shiny.shinyapp.$socket;
  if (sk && sk.readyState === 3) return 'socket';
} catch (e) {}
return '';
"""

JS_CLICK_RELOAD = """
for (const id of ['ss-reload-link', 'shiny-reload-link']) {
  const el = document.getElementById(id);
  if (el) { el.click(); return 'link:' + id; }
}
for (const id of ['ss-connect-dialog', 'shiny-disconnected-overlay']) {
  const box = document.getElementById(id);
  const a = box && box.querySelector('a, button');
  if (a) { a.click(); return 'anchor:' + id; }
}
// the reload control inside whatever element carries the disconnect message
const rx = /disconnected from the server/i;
for (const el of document.querySelectorAll('div, p, span, section')) {
  if (el.children.length > 3) continue;
  if (!rx.test(el.textContent || '')) continue;
  const a = el.querySelector('a, button');
  if (a) { a.click(); return 'msg-anchor'; }
}
const a = Array.from(document.querySelectorAll('a, button')).find(
  e => (e.innerText || '').trim().toLowerCase() === 'reload');
if (a) { a.click(); return 'text'; }
return 'none';
"""

JS_INSERTIONS = r"""
const rx = /selected\s+([0-9,]+)\s+insertion/i;
for (const el of document.querySelectorAll('div, p, span, label, b, strong')) {
  if (el.children.length > 2) continue;
  const m = rx.exec(el.textContent || '');
  if (m) return parseInt(m[1].replace(/,/g, ''), 10);
}
return -1;
"""

JS_READ_TOKEN = """
const el = document.querySelector('input[type=password]');
return el ? (el.value || '') : '';
"""

JS_ON_ACTIVE_PANE = """
const el = arguments[0];
let node = el;
while (node && node.closest) {
  const pane = node.closest('.tab-pane');
  if (!pane) break;
  if (!pane.classList.contains('active')) return false;
  node = pane.parentElement;
}
if (el.offsetParent !== null) return true;
if (getComputedStyle(el).position === 'fixed') return true;
// radios and checkboxes are often visually replaced by a styled label
const lab = el.closest('label');
if (lab && lab.offsetParent !== null) return true;
const wrap = el.parentElement;
return !!(wrap && wrap.offsetParent !== null);
"""

JS_DUMP_DOM = """
const out = [];
document.querySelectorAll(
  'input, select, button, a[data-toggle], a[data-bs-toggle], .btn, [role=tab]'
).forEach(el => {
  const r = el.getBoundingClientRect();
  out.push({
    tag: el.tagName.toLowerCase(), type: el.getAttribute('type'),
    id: el.id || null, name: el.getAttribute('name'),
    cls: el.className && el.className.toString().slice(0, 120),
    text: (el.innerText || el.value || '').trim().slice(0, 80),
    label: (el.closest('label') ? el.closest('label').innerText.trim().slice(0, 60) : null),
    visible: !!(r.width && r.height)
  });
});
return out;
"""

# Busy check. Deliberately strict about what counts as a *visible* spinner:
# shinycssloaders and friends leave zero-size / opacity-0 spinner divs in the
# DOM permanently, and treating those as "busy" hangs the run forever.
JS_BUSY = """
function shown(n) {
  const r = n.getBoundingClientRect();
  if (!r.width || !r.height) return false;
  const s = getComputedStyle(n);
  if (s.display === 'none' || s.visibility === 'hidden') return false;
  if (parseFloat(s.opacity || '1') < 0.05) return false;
  return true;
}
if (document.readyState !== 'complete') return 'readyState';
if (document.documentElement.classList.contains('shiny-busy')) return 'shiny-busy';
for (const n of document.querySelectorAll('.recalculating')) {
  if (shown(n)) return 'recalculating';
}
for (const n of document.querySelectorAll('.load-container, .spinner, .shiny-spinner-output-container')) {
  if (shown(n)) return 'spinner';
}
return '';
"""


# --------------------------------------------------------------------------- #
# Browser
# --------------------------------------------------------------------------- #

LOCK_FILES = ("SingletonLock", "SingletonCookie", "SingletonSocket", "lockfile")


def chrome_processes() -> int:
    """How many chrome.exe are running right now (Windows)."""
    try:
        out = subprocess.run(["tasklist", "/FI", "IMAGENAME eq chrome.exe", "/NH"],
                             capture_output=True, text=True, timeout=15).stdout
        return sum(1 for ln in out.splitlines() if "chrome.exe" in ln.lower())
    except (OSError, subprocess.SubprocessError):
        return -1


def clear_stale_locks(profile: Path) -> list[str]:
    """Remove Chrome's singleton lock files. Only safe with no Chrome running."""
    removed = []
    for name in LOCK_FILES:
        f = Path(profile) / name
        try:
            if f.exists() or f.is_symlink():
                f.unlink()
                removed.append(name)
        except OSError:
            pass
    return removed


def _chrome_options(cfg: Config, dl: str, profile: Path | None) -> ChromeOptions:
    o = ChromeOptions()
    o.add_argument("--window-size=1920,1200")
    o.add_argument("--disable-gpu")
    o.add_argument("--no-sandbox")
    o.add_argument("--disable-dev-shm-usage")
    o.add_argument("--disable-blink-features=AutomationControlled")
    if profile:
        Path(profile).mkdir(parents=True, exist_ok=True)
        o.add_argument(f"--user-data-dir={Path(profile).resolve()}")
        o.add_argument("--profile-directory=Default")
    o.add_experimental_option("excludeSwitches", ["enable-automation"])
    o.add_experimental_option("useAutomationExtension", False)
    o.add_experimental_option("prefs", {
        "download.default_directory": dl,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1,
    })
    return o


def build_driver(cfg: Config, download_dir: Path) -> webdriver.Chrome:
    """
    Launch Chrome, recovering from the usual Windows failure: the profile
    directory is still locked by a Chrome that is open (or was killed and left
    a stale SingletonLock behind). "Chrome instance exited" is what that looks
    like, and the raw stack trace says nothing useful.
    """
    from selenium.common.exceptions import SessionNotCreatedException

    download_dir.mkdir(parents=True, exist_ok=True)
    dl = str(download_dir.resolve())

    def say(msg: str) -> None:
        cb = getattr(cfg, "on_log", None)
        print(msg, flush=True)
        if callable(cb):
            try:
                cb(msg)
            except Exception:
                pass

    profile = Path(cfg.profile_dir) if cfg.profile_dir else None
    attempts: list[tuple[str, Path | None]] = [("configured profile", profile)]

    try:
        drv = webdriver.Chrome(options=_chrome_options(cfg, dl, profile))
    except SessionNotCreatedException as first:
        running = chrome_processes()
        say(f"Chrome failed to start ({running if running >= 0 else '?'} "
            f"chrome.exe running).")

        recovered = None
        if profile:
            if running == 0:
                gone = clear_stale_locks(profile)
                if gone:
                    say(f"Cleared stale lock file(s): {', '.join(gone)} - retrying "
                        f"with the same profile so your login is kept.")
                try:
                    recovered = webdriver.Chrome(
                        options=_chrome_options(cfg, dl, profile))
                except SessionNotCreatedException:
                    recovered = None
            else:
                say("That profile is in use by a Chrome that is already open. "
                    "Close every Chrome window started by this tool (or all of "
                    "them) and press '1. Open browser' again.")

            if recovered is None:
                alt = profile.with_name(
                    profile.name + "_alt_" + datetime.now().strftime("%H%M%S"))
                say(f"Falling back to a separate profile: {alt.name} "
                    f"(you will have to sign in again in this window).")
                attempts.append(("fallback profile", alt))
                try:
                    recovered = webdriver.Chrome(
                        options=_chrome_options(cfg, dl, alt))
                except SessionNotCreatedException:
                    recovered = None

        if recovered is None:
            raise RuntimeError(
                "Chrome would not start.\n\n"
                f"chrome.exe processes running: "
                f"{running if running >= 0 else 'unknown'}\n"
                f"Profile tried: {profile}\n\n"
                "Most likely causes, in order:\n"
                "  1. A Chrome using this profile is still open - close it "
                "(check Task Manager for chrome.exe) and try again.\n"
                "  2. Chrome was updated and chromedriver is stale - Selenium "
                "Manager usually fixes this itself; restarting this app forces "
                "a re-check.\n"
                "  3. Corporate policy or antivirus is blocking the launch.\n\n"
                f"Original error: {first.msg if hasattr(first, 'msg') else first}"
            ) from first
        drv = recovered
        say(f"Chrome started on the {attempts[-1][0]}.")

    drv.set_page_load_timeout(cfg.timeout * 2)
    set_download_dir(drv, download_dir)
    return drv


def set_download_dir(drv, folder: Path) -> None:
    """Point Chrome's downloads at `folder` (works after launch, via CDP)."""
    folder.mkdir(parents=True, exist_ok=True)
    for cmd in ("Browser.setDownloadBehavior", "Page.setDownloadBehavior"):
        try:
            drv.execute_cdp_cmd(cmd, {"behavior": "allow",
                                      "downloadPath": str(folder.resolve())})
        except WebDriverException:
            pass


# --------------------------------------------------------------------------- #
# Bot -- thin wrapper around the driver
# --------------------------------------------------------------------------- #

class Cancelled(Exception):
    pass


def iso_to_page(iso: str, fmt: str) -> str:
    """
    Render an ISO date the way this page's datepicker writes them.

    A picker set to dd/mm/yyyy silently rejects '2022-03-13' and snaps back to
    what it had, which reads exactly like "the filter is stuck".
    """
    iso = (iso or "").strip()
    if not iso:
        return ""
    try:
        d = datetime.strptime(iso[:10], "%Y-%m-%d")
    except ValueError:
        return iso
    f = (fmt or "").strip()
    if not f or f.lower() == "yyyy-mm-dd":
        return d.strftime("%Y-%m-%d")
    # bootstrap-datepicker spells month/weekday NAMES with capitals (M, MM,
    # D, DD). Those are not ours to render, so fall back to ISO.
    if re.search(r"[MD]", f):
        return d.strftime("%Y-%m-%d")
    f = f.lower()
    out = f
    for tok, val in (("yyyy", f"{d.year:04d}"), ("yy", f"{d.year % 100:02d}"),
                     ("mm", f"{d.month:02d}"), ("dd", f"{d.day:02d}")):
        out = out.replace(tok, val)
    out = re.sub(r"(?<!\d)m(?!\d)", str(d.month), out)
    out = re.sub(r"(?<!\d)d(?!\d)", str(d.day), out)
    if re.search(r"[A-Za-z]", out):        # month names and the like: no guess
        return d.strftime("%Y-%m-%d")
    return out


def page_to_iso(val: str, fmt: str = "") -> str:
    """Whatever the box shows -> YYYY-MM-DD, so the rest of the code is sane."""
    v = (val or "").strip()
    if not v:
        return ""
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", v):
        return v
    m = re.fullmatch(r"(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{4})", v)
    if m:
        a, b, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        f = (fmt or "").strip().lower()
        day, mon = a, b
        if f.startswith("m"):            # mm/dd/yyyy
            day, mon = b, a
        elif a > 12:
            day, mon = a, b
        elif b > 12:
            day, mon = b, a
        if 1 <= mon <= 12 and 1 <= day <= 31:
            return f"{y:04d}-{mon:02d}-{day:02d}"
        return v
    for f in ("%d %b %Y", "%d %B %Y", "%b %d, %Y", "%B %d, %Y", "%Y/%m/%d"):
        try:
            return datetime.strptime(v, f).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return v


class Bot:
    def __init__(self, drv, cfg: Config):
        self.d = drv
        self.cfg = cfg
        self.log_lines: list[str] = []
        self.log_file: Path | None = None      # set once the run folder exists
        self._last_disc = ""
        self.skipped_feeds: list[str] = []
        self.was_cancelled = False

    def set_log_file(self, path: Path) -> None:
        """Mirror the log to disk from here on, flushing every line."""
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            with open(path, "a", encoding="utf-8") as fh:
                fh.write("\n".join(self.log_lines) + "\n")
            self.log_file = path
            self.log(f"log file: {path}")
        except OSError as exc:
            self.log(f"note: could not open a log file ({exc})")

    # -- logging --------------------------------------------------------- #
    def log(self, msg: str) -> None:
        line = f"[{datetime.now():%H:%M:%S}] {msg}"
        print(line, flush=True)
        self.log_lines.append(line)
        if self.log_file is not None:
            try:                       # flushed per line: survives a hard crash
                with open(self.log_file, "a", encoding="utf-8") as fh:
                    fh.write(line + "\n")
            except OSError:
                self.log_file = None
        cb = getattr(self.cfg, "on_log", None)
        if callable(cb):
            try:
                cb(line)
            except Exception:
                pass

    def step(self, msg: str) -> None:
        """Announce a stage BEFORE doing it, so a stall is always visible."""
        self.log(f">> {msg}")

    def progress(self, **kw) -> None:
        cb = getattr(self.cfg, "on_progress", None)
        if callable(cb):
            try:
                cb(kw)
            except Exception:
                pass

    def check_cancel(self) -> None:
        ev = getattr(self.cfg, "cancel", None)
        if ev is not None and ev.is_set():
            raise Cancelled("cancelled by user")

    def visible_tabs(self) -> list[str]:
        try:
            return self.d.execute_script(
                "return Array.from(document.querySelectorAll("
                "'a[data-toggle=tab],a[data-bs-toggle=tab],[role=tab]'))"
                ".map(e => e.innerText.trim()).filter(Boolean);") or []
        except WebDriverException:
            return []

    # -- finding --------------------------------------------------------- #
    def disconnected(self) -> bool:
        try:
            why = self.d.execute_script(JS_DISCONNECTED) or ""
        except WebDriverException:
            return False
        if why and why != self._last_disc:
            self.log(f"  (disconnect detected via {why})")
        self._last_disc = why
        return bool(why)

    def insertions(self) -> int:
        """'Selected N insertion(s)' -- 0 means this feed has nothing to plot."""
        try:
            return int(self.d.execute_script(JS_INSERTIONS))
        except (WebDriverException, TypeError, ValueError):
            return -1

    def on_active_pane(self, el) -> bool:
        """Is this element on the tab we are actually looking at?"""
        try:
            return bool(self.d.execute_script(JS_ON_ACTIVE_PANE, el))
        except WebDriverException:
            return True

    def find_any(self, locators, timeout: int, required: bool, what: str):
        start = time.time()
        deadline = start + timeout
        last = None
        announced = False
        while time.time() < deadline:
            self.check_cancel()
            for by, sel in locators:
                try:
                    for el in self.d.find_elements(by, sel):
                        # never touch a control belonging to a hidden tab
                        if not self.on_active_pane(el):
                            continue
                        if el.is_displayed():
                            return el
                        if el.tag_name == "input" and \
                                el.get_attribute("type") in ("checkbox", "radio"):
                            return el
                except (WebDriverException, StaleElementReferenceException) as e:
                    last = e
            if not announced and time.time() - start > 5:
                announced = True
                self.log(f"  still looking for '{what}' "
                         f"({len(locators)} locators, {timeout}s budget)...")
            time.sleep(0.3)
        if required:
            raise TimeoutException(
                f"Could not find '{what}' after {timeout}s. Tried "
                f"{len(locators)} locator(s). Tabs on page: {self.visible_tabs()}. "
                f"Last error: {last}")
        return None

    def find(self, key: str, timeout: int | None = None, required: bool = True):
        return self.find_any(SELECTORS[key], timeout or self.cfg.timeout,
                             required, key)

    def click(self, el_or_key, name: str = "") -> None:
        el = self.find(el_or_key) if isinstance(el_or_key, str) else el_or_key
        what = el_or_key if isinstance(el_or_key, str) else (name or "element")
        self.d.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        time.sleep(0.15)
        try:
            el.click()
        except (ElementClickInterceptedException, WebDriverException):
            self.d.execute_script("arguments[0].click();", el)
        self.log(f"clicked {what}")

    def set_checkbox(self, key: str, want: bool) -> None:
        el = self.find(key)
        if el.is_selected() != want:
            self.click(el, key)
        else:
            self.log(f"{key} already {'ticked' if want else 'unticked'}")
        self.d.execute_script(
            "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", el)

    def set_checkbox_labelled(self, label: str, want: bool) -> bool:
        """Tick/untick by visible label, then read it back to prove it stuck."""
        el = self.find_any(checkbox_locators(label), 20, False, f"'{label}'")
        if el is None:
            self.log(f"WARN: checkbox '{label}' not found on this tab - skipped")
            return False
        for attempt in (1, 2):
            if el.is_selected() == want:
                break
            self.click(el, f"'{label}'")
            self.d.execute_script(
                "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", el)
            time.sleep(0.3)
        got = el.is_selected()
        if got != want:
            self.log(f"WARN: '{label}' is {got}, wanted {want} - the plot may "
                     f"not match the requested options")
            return False
        self.log(f"'{label}' = {got}")
        return True

    def set_marker(self, label: str) -> bool:
        """Select a Marker type radio and confirm the page agrees."""
        el = self.find_any(radio_locators(label), 20, False, f"marker '{label}'")
        if el is None:
            self.log(f"WARN: marker radio '{label}' not found on this tab")
            return False
        for attempt in (1, 2):
            try:
                if el.tag_name == "input" and el.is_selected():
                    break
            except WebDriverException:
                pass
            self.click(el, f"marker '{label}'")
            self.d.execute_script(
                "arguments[0].dispatchEvent(new Event('change',{bubbles:true}));", el)
            time.sleep(0.3)
        try:
            ok = el.tag_name != "input" or el.is_selected()
        except WebDriverException:
            ok = True
        # report what the page actually has selected, whatever we clicked
        try:
            chosen = self.d.execute_script("""
                const el = arguments[0];
                if (el.tagName.toLowerCase() !== 'input') return '';
                const group = document.getElementsByName(el.name || '');
                for (const r of group) {
                  if (r.checked) {
                    const lab = r.closest('label');
                    return (lab ? lab.innerText : r.value).trim();
                  }
                }
                return '';
            """, el)
        except WebDriverException:
            chosen = ""
        self.log(f"marker type = {chosen or label}"
                 + ("" if ok else f" (WARN: wanted '{label}')"))
        return bool(ok)

    # -- date range ------------------------------------------------------ #
    def read_date_range(self) -> dict:
        """
        What the app itself has in its Date Range box, right now.

        Returns {} when there is no such control on the page. 'from_iso' and
        'to_iso' are normalised; 'from'/'to' are what the box literally shows.
        """
        try:
            got = self.d.execute_script(JS_DATE_RANGE)
        except WebDriverException:
            got = None
        if not isinstance(got, dict):
            return {}
        fmt = (got.get("fmt") or "").strip()
        got["from_iso"] = page_to_iso(got.get("from", ""), fmt)
        got["to_iso"] = page_to_iso(got.get("to", ""), fmt)
        return got

    def date_inputs(self):
        """The two Date Range <input>s, or (None, None)."""
        try:
            got = self.d.execute_script(JS_DATE_INPUTS)
        except WebDriverException:
            got = None
        if isinstance(got, list) and len(got) == 2 and all(got):
            return got[0], got[1]
        return None, None

    def set_date(self, key: str, value: str) -> bool:
        """
        Push one end of the Date Range in, then read the page back.

        Returns True only when the page kept it. A datepicker that rejects the
        value (wrong format, or outside the data) reverts silently, and that is
        worth a warning rather than a report quoting dates nobody applied.
        """
        a, z = self.date_inputs()
        el = a if key == "date_from" else z
        if el is None:
            el = self.find(key, required=False)
        if el is None:
            self.log(f"WARN: no '{key}' field on this tab - "
                     f"leaving the app's own range alone")
            return False

        info = self.read_date_range()
        fmt = (info.get("fmt") or "").strip()
        typed = iso_to_page(value, fmt) or value
        try:
            current = (el.get_attribute("value") or "").strip()
        except WebDriverException:
            current = ""
        if current == typed:
            self.log(f"{key} already {typed} - left alone")
            return True
        self.log(f"{key}: '{current}' -> '{typed}'"
                 + (f" (page format {fmt})" if fmt and fmt != "yyyy-mm-dd" else ""))
        self.d.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
        try:
            el.click()
            el.send_keys(Keys.CONTROL, "a")
            el.send_keys(Keys.DELETE)
            el.send_keys(typed)
            time.sleep(0.25)
            el.send_keys(Keys.ESCAPE)
        except WebDriverException:
            pass
        # belt and braces for bootstrap-datepicker / Shiny bindings
        try:
            self.d.execute_script("""
                const el = arguments[0], v = arguments[1];
                el.value = v;
                el.dispatchEvent(new Event('input',  {bubbles:true}));
                el.dispatchEvent(new Event('change', {bubbles:true}));
                if (window.jQuery) {
                    jQuery(el).trigger('change');
                    try { jQuery(el).datepicker('update', v); } catch (e) {}
                }
            """, el, typed)
            self.d.execute_script("document.body.click();")
        except WebDriverException:
            pass

        # read the page again rather than the element we typed into: Shiny
        # often replaces the input, which would leave us reading a ghost
        after = self.read_date_range()
        got = (after.get("from_iso") if key == "date_from"
               else after.get("to_iso")) or ""
        if got == value:
            self.log(f"{key} = {value}")
            return True
        bounds = ""
        if after.get("min") or after.get("max"):
            bounds = (f"; the app allows {after.get('min') or '?'} .. "
                      f"{after.get('max') or '?'}")
        self.log(f"WARN: {key} would not take '{value}' - the box now reads "
                 f"'{got or '?'}'{bounds}")
        return False

    # -- waiting --------------------------------------------------------- #
    def wait_idle(self, quiet: float = 1.0, timeout: int | None = None) -> None:
        """
        Best-effort settle. Never blocks the pipeline: if the page keeps
        claiming to be busy we say so and carry on, because a permanently
        visible spinner div is far more likely than a 20-second computation.
        """
        limit = timeout or self.cfg.idle_timeout
        start = time.time()
        deadline = start + limit
        stable = None
        why = ""
        while time.time() < deadline:
            self.check_cancel()
            try:
                why = self.d.execute_script(JS_BUSY) or ""
            except WebDriverException:
                why = ""
            if self.disconnected():
                self.log("  (Shiny reports disconnected while waiting)")
                return
            if why:
                stable = None
            else:
                stable = stable or time.time()
                if time.time() - stable >= quiet:
                    took = time.time() - start
                    if took > 3:
                        self.log(f"  (page settled after {took:.1f}s)")
                    return
            time.sleep(0.2)
        self.log(f"  (page still reports '{why}' after {limit}s - continuing)")

    def wait_plots(self, timeout: int | None = None,
                   settle: float = 1.0) -> int:
        """Wait until the number of rendered charts stops changing for `settle`
        seconds -- Media renders dozens progressively, so a short window would
        return while charts are still appearing."""
        deadline = time.time() + (timeout or self.cfg.plot_timeout)
        need = max(2, int(settle / 0.25))
        last, stable = 0, 0
        announced = 0
        while time.time() < deadline:
            self.check_cancel()
            res = self.d.execute_script(JS_COUNT_PLOTS) or {}
            n, ready = res.get("n", 0), res.get("ready", False)
            if n != announced and n:
                announced = n
                self.log(f"  ...{n} chart(s) so far")
            if ready and n == last and n > 0:
                stable += 1
                if stable >= need:
                    self.log(f"{n} chart(s) rendered")
                    return n
            else:
                stable = 0
            last = n
            time.sleep(0.25)
        raise TimeoutException(
            "No Plotly charts appeared. Did 'Go Plot' actually run?")

    def dump_dom(self, path: Path) -> None:
        try:
            data = self.d.execute_script(JS_DUMP_DOM)
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(data, indent=2), encoding="utf-8")
            self.log(f"DOM dump -> {path.name} ({len(data)} elements)")
        except WebDriverException as e:
            self.log(f"DOM dump failed: {e}")


# --------------------------------------------------------------------------- #
# Plot handling
# --------------------------------------------------------------------------- #

def resolve_plot_index(bot: Bot, title: str, gi: int) -> int:
    """Charts can be re-rendered mid-run; find this one by its title again."""
    if not title:
        return gi
    try:
        found = bot.d.execute_script(JS_FIND_BY_TITLE, title)
    except WebDriverException:
        return gi
    if isinstance(found, int) and found >= 0:
        if found != gi:
            bot.log(f"  note: '{title}' moved from index {gi} to {found}")
        return found
    return gi


def set_visible_traces(bot: Bot, gi: int, keep: list[str]) -> bool:
    """
    Show only `keep`; grey out the rest. Returns False when nothing matched --
    the chart is then left exactly as it was, because hiding every series would
    export a blank PNG and produce observations that do not match the picture.
    """
    try:
        res = bot.d.execute_script(JS_SET_VISIBILITY, gi, keep)
    except JavascriptException as e:
        bot.log(f"  restyle failed ({e}); trying the legend")
        return _legend_click_fallback(bot, gi, keep)

    if isinstance(res, dict) and res.get("ok"):
        bot.log(f"  chart {gi}: showing only {keep} "
                f"({res.get('matched')}/{res.get('total')} series)")
        time.sleep(0.25)
        return True

    reason = res.get("reason") if isinstance(res, dict) else res
    if reason == "no-match":
        bot.log(f"  WARN: chart {gi} has no series called {keep}; it carries "
                f"{res.get('have')}. Left untouched.")
        return False
    bot.log(f"  WARN: restyle said '{reason}' for chart {gi}; trying the legend")
    return _legend_click_fallback(bot, gi, keep)


def _legend_click_fallback(bot: Bot, gi: int, keep: list[str]) -> bool:
    keepn = {norm(k) for k in keep}
    gd = bot.d.execute_script(JS_GET_PLOT, gi)
    if gd is None:
        bot.log(f"  WARN: chart {gi} is not on the visible tab")
        return False
    try:
        info = bot.d.execute_script(JS_TRACE_INFO)[gi]["traces"]
    except (WebDriverException, IndexError, KeyError, TypeError):
        return False
    if not any(norm(t["name"]) in keepn for t in info):
        bot.log(f"  WARN: chart {gi} has no series called {keep} - left untouched")
        return False
    items = gd.find_elements(C, "g.legend g.traces")
    for i, item in enumerate(items):
        if i >= len(info):
            break
        if (norm(info[i]["name"]) in keepn) != (info[i]["visible"] is True):
            bot.d.execute_script("arguments[0].scrollIntoView({block:'center'});",
                                 item)
            item.click()
            time.sleep(0.4)
    bot.log(f"  chart {gi}: legend toggled to {keep}")
    return True


def download_plot_png(bot: Bot, gi: int, filename: str, folder: Path) -> Path:
    """
    Plotly's own 'Download plot as png' export -- same code path as the camera
    icon in the modebar, but with a filename we control.  Falls back to actually
    clicking the camera and renaming whatever lands in the folder.
    """
    stamp = time.time()
    status = "?"
    try:
        status = bot.d.execute_async_script(
            JS_DOWNLOAD_IMAGE, gi, filename, bot.cfg.img_width, bot.cfg.img_height)
    except (JavascriptException, TimeoutException) as e:
        status = f"err:{e}"

    if not str(status).startswith("ok"):
        bot.log(f"  downloadImage said '{status}' -> clicking the camera icon")
        gd = bot.d.execute_script(JS_GET_PLOT, gi)
        if gd is None:
            raise RuntimeError(f"chart {gi} is not on the visible tab")
        bot.d.execute_script("arguments[0].scrollIntoView({block:'center'});", gd)
        cam = gd.find_elements(
            C, "a[data-title*='Download plot'], a.modebar-btn[data-title*='png']") \
            or gd.find_elements(C, ".modebar-btn")[:1]
        if not cam:
            raise RuntimeError(f"No download button on chart {gi}")
        bot.d.execute_script("arguments[0].click();", cam[0])

    target = folder / f"{filename}.png"
    deadline = time.time() + 60
    while time.time() < deadline:
        bot.check_cancel()
        if target.exists() and target.stat().st_size:
            return target
        fresh = [p for p in folder.glob("*.png")
                 if p != target and p.stat().st_mtime >= stamp - 1]
        if fresh:
            got = max(fresh, key=lambda p: p.stat().st_mtime)
            time.sleep(0.15)                     # let the write finish
            try:
                shutil.move(str(got), str(target))
            except OSError:
                time.sleep(0.4)
                shutil.move(str(got), str(target))
            return target
        time.sleep(0.1)
    raise TimeoutException(f"PNG '{filename}' never arrived in {folder}")


# --------------------------------------------------------------------------- #
# Observations from the real series
# --------------------------------------------------------------------------- #

def _f(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def analyze_series(x: list, y: list, name: str) -> list[str]:
    pairs = [(str(a), _f(b)) for a, b in zip(x, y)]
    vals = [v for _, v in pairs if v is not None]
    if not vals:
        return [f"{name}: no numeric data in the selected range."]

    n = len(vals)
    lo, hi = min(vals), max(vals)
    mean = sum(vals) / n
    srt = sorted(vals)
    med = srt[n // 2] if n % 2 else (srt[n // 2 - 1] + srt[n // 2]) / 2
    sd = (sum((v - mean) ** 2 for v in vals) / n) ** 0.5
    mad = sorted(abs(v - med) for v in vals)[n // 2] or 1e-9

    obs = [f"Coverage: {n} points from {pairs[0][0][:10]} to {pairs[-1][0][:10]}. "
           f"Mean {mean:,.1f}, median {med:,.1f}, SD {sd:,.1f} "
           f"(CV {0 if mean == 0 else sd / abs(mean):.2f}), range {lo:,.1f}-{hi:,.1f}."]

    scored = [(d, v, (v - med) / (1.4826 * mad)) for d, v in pairs if v is not None]
    peaks = sorted((s for s in scored if s[2] >= 3), key=lambda s: -s[2])[:5]
    dips = sorted((s for s in scored if s[2] <= -3), key=lambda s: s[2])[:5]
    if peaks:
        obs.append("Peaks (robust z >= 3): " + "; ".join(
            f"{d[:10]} = {v:,.0f} ({z:+.1f} sigma, "
            f"{0 if med == 0 else v / med:.1f}x median)" for d, v, z in peaks) + ".")
    else:
        obs.append("No point exceeds a robust z-score of 3 - no isolated spikes.")
    if dips:
        obs.append("Dips (robust z <= -3): " + "; ".join(
            f"{d[:10]} = {v:,.0f} ({z:+.1f} sigma)" for d, v, z in dips) + ".")

    # recurring seasonality: same month, across at least two different years
    top = sorted(scored, key=lambda s: -s[1])[:max(3, n // 10)]
    months: dict[str, int] = {}
    years: dict[str, set] = {}
    for d, _, _ in top:
        m, yr = d[5:7], d[0:4]
        if re.fullmatch(r"\d{2}", m) and 1 <= int(m) <= 12:
            months[m] = months.get(m, 0) + 1
            years.setdefault(m, set()).add(yr)
    if months:
        dom, cnt = max(months.items(), key=lambda kv: kv[1])
        if cnt >= 2 and cnt / len(top) >= 0.4 and len(years[dom]) >= 2:
            obs.append(f"Seasonality: {cnt} of the top {len(top)} values fall in "
                       f"{datetime(2000, int(dom), 1):%B}, across {len(years[dom])} "
                       f"different years - recurring annual peak.")

    runs, run, start = [], 0, None
    for d, v in pairs:
        if v is not None and abs(v) < 1e-12:
            run += 1
            start = start or d
        else:
            if run >= 3:
                runs.append((start, d, run))
            run, start = 0, None
    if run >= 3:
        runs.append((start, pairs[-1][0], run))
    if runs:
        obs.append("Zero/flat stretches (>=3 consecutive points): " + "; ".join(
            f"{a[:10]} -> {b[:10]} ({k} pts)" for a, b, k in runs[:4]) + ".")

    gaps = sum(1 for _, v in pairs if v is None)
    if gaps:
        obs.append(f"{gaps} missing point(s) - {gaps / len(pairs):.0%} of the range.")

    third = max(1, n // 3)
    a = sum(vals[:third]) / third
    b = sum(vals[-third:]) / third
    if a:
        delta = (b - a) / abs(a)
        if abs(delta) >= 0.10:
            obs.append(f"Trend: last third averages {b:,.0f} vs {a:,.0f} in the first "
                       f"third ({delta:+.0%}) - "
                       f"{'upward' if delta > 0 else 'downward'} drift.")
        else:
            obs.append(f"Trend: broadly flat ({delta:+.0%} first third to last third).")

    steps = [(pairs[i][0], vals[i] - vals[i - 1]) for i in range(1, len(vals))]
    if steps:
        d, jump = max(steps, key=lambda s: abs(s[1]))
        if sd and abs(jump) > 3 * sd:
            obs.append(f"Largest single-period move: {jump:+,.0f} at {d[:10]} "
                       f"({abs(jump) / sd:.1f}x SD) - check for a data break.")

    if lo < 0:
        obs.append(f"Negative values present (min {lo:,.1f}) - unexpected for a "
                   f"count/volume metric; worth validating.")
    return obs


# --------------------------------------------------------------------------- #
# House-style narrative -- matches the wording used in the sanity-check reports:
#   "During the October 2025 to December 2025 refit period, Physical Signatures
#    showed a notable increase and reached its highest levels during the refit
#    period. Can you please confirm what drove this uplift?"
# --------------------------------------------------------------------------- #

def month_year(d: str) -> str:
    try:
        return datetime.strptime(d[:10], "%Y-%m-%d").strftime("%B %Y")
    except ValueError:
        return d[:10]


def _episodes(flagged: list[int], max_gap: int = 1) -> list[list[int]]:
    """Group nearby flagged indices into episodes, so a run of high weeks
    becomes one 'October 2025 to December 2025' window rather than 6 bullets."""
    groups: list[list[int]] = []
    for i in flagged:
        if groups and i - groups[-1][-1] <= max_gap + 1:
            groups[-1].append(i)
        else:
            groups.append([i])
    return groups


def _window(pairs, g) -> str:
    a, b = month_year(pairs[g[0]][0]), month_year(pairs[g[-1]][0])
    return a if a == b else f"{a} to {b}"


def narrative_observations(x: list, y: list, metric: str = "") -> list[str]:
    """
    Terse observation fragments in the BT house style -- the box already asks
    "Are the following observations correct?", so each line is a statement of
    what the chart shows, not a demand for an explanation. The UID is printed
    above them, so the metric name is not repeated on every line.

    Real examples from the BT report this mirrors:
        "Peak on July 04th 2025."
        "No data for target period"
        "Peaks and dips in whole of July"
    """
    pairs = [(str(a), _f(b)) for a, b in zip(x, y)]
    pairs = [(d, v) for d, v in pairs if v is not None]
    if len(pairs) < 3:
        return ["Not enough data in the selected range to comment on the trend."]

    vals = [v for _, v in pairs]
    n = len(vals)
    srt = sorted(vals)
    med = srt[n // 2] if n % 2 else (srt[n // 2 - 1] + srt[n // 2]) / 2
    mean = sum(vals) / n
    mad = sorted(abs(v - med) for v in vals)[n // 2] or 1e-9
    z = [(v - med) / (1.4826 * mad) for v in vals]
    hi_i = max(range(n), key=lambda i: vals[i])
    out: list[str] = []

    def amount(v: float) -> str:
        return f"{v:,.0f}" if abs(v) >= 100 else f"{v:,.2f}".rstrip("0").rstrip(".")

    # ---- overall shape ------------------------------------------------- #
    zero_n = sum(1 for v in vals if abs(v) < 1e-12)
    third = max(1, n // 3)
    a3 = sum(vals[:third]) / third
    b3 = sum(vals[-third:]) / third
    drift = (b3 - a3) / abs(a3) if a3 else 0.0

    if zero_n / n >= 0.4:
        active = [v for v in vals if abs(v) > 1e-12]
        out.append(f"Sparse series - {zero_n} of {n} periods ({zero_n / n:.0%}) "
                   f"are zero, averaging "
                   f"{amount(sum(active) / len(active)) if active else '0'} "
                   f"when active.")
    elif abs(drift) < 0.10:
        out.append(f"Trend looks stable across the period, averaging "
                   f"{amount(mean)}.")
    else:
        out.append(f"Overall {'increase' if drift > 0 else 'decline'} across the "
                   f"period - average moves from {amount(a3)} to {amount(b3)} "
                   f"({drift:+.0%}).")

    # ---- recurring seasonality ------------------------------------------ #
    months: dict[str, set] = {}
    for i in sorted(range(n), key=lambda i: -vals[i])[:max(3, n // 10)]:
        mm = pairs[i][0][5:7]
        if re.fullmatch(r"\d{2}", mm) and 1 <= int(mm) <= 12:
            months.setdefault(mm, set()).add(pairs[i][0][0:4])
    season = next(((m, yrs) for m, yrs in
                   sorted(months.items(), key=lambda kv: -len(kv[1]))
                   if len(yrs) >= 2), None)
    if season:
        mname = datetime(2000, int(season[0]), 1).strftime("%B")
        out.append(f"Recurring seasonal peak around {mname} "
                   f"({', '.join(sorted(season[1]))}).")

    # ---- peaks ----------------------------------------------------------- #
    peaks = _episodes([i for i in range(n) if z[i] >= 3])[:3]
    if peaks:
        bits = []
        for g in peaks:
            top = max(g, key=lambda i: vals[i])
            r = (vals[top] / med) if med else 0
            ratio = f", {r:.1f}x median" if r >= 1.15 else ""
            bits.append(f"{pairs[top][0][:10]} ({amount(vals[top])}{ratio})")
        highest = " - the highest in the period" if any(
            max(g, key=lambda i: vals[i]) == hi_i for g in peaks) else ""
        out.append(f"Peak{'s' if len(bits) > 1 else ''} on "
                   f"{'; '.join(bits)}{highest}.")

    # ---- dips ------------------------------------------------------------ #
    dips = _episodes([i for i in range(n)
                      if z[i] <= -3 and abs(vals[i]) > 1e-12])[:3]
    if dips:
        bits = []
        for g in dips:
            low = min(g, key=lambda i: vals[i])
            bits.append(f"{pairs[low][0][:10]} ({amount(vals[low])})")
        out.append(f"Dip{'s' if len(bits) > 1 else ''} on {'; '.join(bits)} "
                   f"against a median of {amount(med)}.")

    # ---- zero stretches --------------------------------------------------- #
    for g in _episodes([i for i in range(n) if abs(vals[i]) < 1e-12], max_gap=0):
        if len(g) >= 3:
            span = _window(pairs, g).replace(" to ", " and ")
            out.append(f"No data - zero for {len(g)} consecutive periods between "
                       f"{span}.")

    # ---- step change ------------------------------------------------------ #
    steps = [(i, vals[i] - vals[i - 1]) for i in range(1, n)]
    if steps:
        i, jump = max(steps, key=lambda s: abs(s[1]))
        sd = (sum((v - mean) ** 2 for v in vals) / n) ** 0.5
        before_v = vals[max(0, i - 8):i]
        after_v = vals[i:i + 8]
        before = sum(before_v) / max(1, len(before_v))
        after = sum(after_v) / max(1, len(after_v))
        if sd and abs(jump) > 4 * sd and before and                 abs(after - before) > 0.5 * abs(before):
            out.append(f"Step change at {pairs[i][0][:10]} - average moves from "
                       f"{amount(before)} to {amount(after)}.")

    if min(vals) < 0:
        out.append(f"Contains negative values (minimum {amount(min(vals))}).")

    return out


# --------------------------------------------------------------------------- #
# Word report
#
# Modelled on the BT UK / Telepass Data Sanity Check reports:
#
#   <title paragraph>
#   [ 9x2 info table  - Organization / Brand / Model / slug / UIDs / ... ]
#   Contents                              <- real TOC field, Word fills it in
#   HEADING 1: OBJECTIVE                  <- one per Data App tab
#     [ 4x2 box ] DATA TYPE / Data delivered by / Question / Answer
#     <chart image(s) for that box>
#     [ 4x2 box ] ...
#   HEADING 1: NON-MEDIA DATA
#     ...
#
# Box formatting copied from the BT report: 1.34in + 4.91in columns, Arial,
# bold 10pt labels, single black borders, and per-row label shading.
# --------------------------------------------------------------------------- #

BOX_ROWS = ("DATA TYPE", "Data delivered by", "Question", "Answer")
BOX_SHADES = ("EFEFEF", "E7E6E6", "F4CCCC", "D9EAD3")
BOX_COL_W = (1.34, 4.91)

INFO_ROWS = ("Organization name:", "Brand name:", "Model name:", "Model slug:",
             "UIDs:", "Required action:", "Period impacted:",
             "Data availability:", "Notes:")
INFO_COL_W = (1.74, 4.53)
INFO_SHADE = "EFEFEF"

QUESTION_LEAD = "Are the following observations correct?"
BODY_FONT = "Arial"
BODY_PT = 10


def _qn(tag: str):
    from docx.oxml.ns import qn
    return qn(tag)


def _shade(cell, fill: str) -> None:
    from docx.oxml import OxmlElement
    el = OxmlElement("w:shd")
    el.set(_qn("w:val"), "clear")
    el.set(_qn("w:color"), "auto")
    el.set(_qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(el)


def _box_borders(cell) -> None:
    from docx.oxml import OxmlElement
    borders = OxmlElement("w:tcBorders")
    for side in ("top", "left", "bottom", "right"):
        e = OxmlElement(f"w:{side}")
        e.set(_qn("w:val"), "single")
        e.set(_qn("w:color"), "000000")
        e.set(_qn("w:sz"), "8")
        e.set(_qn("w:space"), "0")
        borders.append(e)
    cell._tc.get_or_add_tcPr().append(borders)


def _write_cell(cell, lines, bold: bool = False, size: int = BODY_PT) -> None:
    """Replace a cell's contents with `lines`, one paragraph each."""
    from docx.shared import Pt
    lines = [l for l in (lines or [""])] or [""]
    for extra in list(cell.paragraphs)[1:]:
        extra._element.getparent().remove(extra._element)
    first = cell.paragraphs[0]
    for r in list(first.runs):
        r._element.getparent().remove(r._element)

    def fill(par, text):
        par.paragraph_format.space_after = Pt(2)
        run = par.add_run(text)
        run.font.name = BODY_FONT
        run.font.size = Pt(size)
        run.bold = bold
        return run

    fill(first, lines[0])
    for line in lines[1:]:
        fill(cell.add_paragraph(), line)


def _append_block(doc, el) -> None:
    """Insert a body-level element before w:sectPr, which must stay last."""
    body = doc.element.body
    sect = body.find(_qn("w:sectPr"))
    if sect is not None:
        sect.addprevious(el)
    else:
        body.append(el)


def add_info_table(doc, info: dict) -> None:
    from docx.shared import Inches
    t = doc.add_table(rows=len(INFO_ROWS), cols=2)
    t.autofit = False
    for i, label in enumerate(INFO_ROWS):
        row = t.rows[i]
        for ci, w in enumerate(INFO_COL_W):
            row.cells[ci].width = Inches(w)
            _box_borders(row.cells[ci])
        _shade(row.cells[0], INFO_SHADE)
        _write_cell(row.cells[0], [label], bold=True)
        _write_cell(row.cells[1], [str(info.get(label, ""))])


def add_box(doc, data_type: str, delivered_by: str, question_lines: list[str],
            answer: str = "") -> None:
    from docx.shared import Inches
    values = [data_type, delivered_by, question_lines, answer]
    t = doc.add_table(rows=len(BOX_ROWS), cols=2)
    t.autofit = False
    for i, label in enumerate(BOX_ROWS):
        row = t.rows[i]
        for ci, w in enumerate(BOX_COL_W):
            row.cells[ci].width = Inches(w)
            _box_borders(row.cells[ci])
        _shade(row.cells[0], BOX_SHADES[i])
        _write_cell(row.cells[0], [label], bold=True)
        v = values[i]
        _write_cell(row.cells[1], v if isinstance(v, list) else [str(v)])


def add_toc(doc) -> None:
    """A real TOC field. Word fills it in on open (see mark_fields_dirty)."""
    from docx.oxml import OxmlElement

    doc.add_paragraph("Contents", style="TOC Heading" if _has_style(
        doc, "TOC Heading") else "Heading 1")
    p = doc.add_paragraph()
    r = p.add_run()._r
    toc = r'TOC \o "1-2" \z \u \h'
    for kind, text in (("begin", None), (None, toc),
                       ("separate", None),
                       (None, "Right-click and choose Update Field to build "
                              "the contents list."),
                       ("end", None)):
        if kind:
            fld = OxmlElement("w:fldChar")
            fld.set(_qn("w:fldCharType"), kind)
            r.append(fld)
        else:
            if text.startswith("TOC"):
                instr = OxmlElement("w:instrText")
                instr.set(_qn("xml:space"), "preserve")
                instr.text = text
                r.append(instr)
            else:
                t = OxmlElement("w:t")
                t.text = text
                r.append(t)


def _has_style(doc, name: str) -> bool:
    try:
        doc.styles[name]
        return True
    except KeyError:
        return False


def mark_fields_dirty(doc) -> None:
    """Tell Word to refresh fields (the TOC) when the document opens."""
    from docx.oxml import OxmlElement
    settings = doc.settings.element
    if settings.find(_qn("w:updateFields")) is None:
        el = OxmlElement("w:updateFields")
        el.set(_qn("w:val"), "true")
        settings.append(el)


def _open_docx(path: Path):
    """
    Open a .docx, copying it to a temp file first. OneDrive keeps files as
    online-only placeholders, and python-docx cannot read those directly --
    the copy forces a hydrate.
    """
    from docx import Document
    try:
        return Document(str(path))
    except Exception:
        tmp = Path(tempfile.gettempdir()) / f"_dataapp_{os.getpid()}_{path.name}"
        shutil.copy2(str(path), str(tmp))
        return Document(str(tmp))


def report_path_for(cfg: Config, run_id: str = "") -> Path:
    """Downloads / data sanity / <project> / <project> ... _YYYYmmdd_HHMMSS.docx

    The timestamp is part of the name on purpose: a report left open in Word
    locks the file, and a fresh name means a new run can never be blocked by
    the previous one.
    """
    project = cfg.project or "project"
    folder = Path(cfg.outdir) if cfg.outdir else DOWNLOADS_ROOT / project
    stamp = run_id or datetime.now().strftime("%Y%m%d_%H%M%S")
    return folder / f"{project} Data Sanity Check Report_{stamp}.docx"


def save_docx_safely(doc, out_path: Path, log) -> Path:
    """
    Save, retrying under a new name if the target is locked (open in Word) or
    otherwise unwritable. Losing a finished run at the last step is not an
    option, so this keeps trying and finally falls back to the temp folder.
    """
    candidates = [out_path]
    for i in range(2, 7):
        candidates.append(out_path.with_name(
            f"{out_path.stem} ({i}){out_path.suffix}"))
    candidates.append(out_path.with_name(
        f"{out_path.stem}_{datetime.now():%H%M%S}{out_path.suffix}"))
    candidates.append(Path(tempfile.gettempdir()) / out_path.name)

    last: Exception | None = None
    for i, cand in enumerate(candidates):
        try:
            cand.parent.mkdir(parents=True, exist_ok=True)
            doc.save(str(cand))
            if i:
                log(f"saved as '{cand.name}' instead - the original was locked")
            return cand
        except (PermissionError, OSError) as exc:
            last = exc
            log(f"could not write '{cand.name}' ({exc.__class__.__name__}: "
                f"{exc}); trying another name")
            time.sleep(1.0)
    raise RuntimeError(
        f"Could not save the report anywhere. Last error: {last}") from last


def build_report(cfg: Config, tab_sections: list[tuple[str, list[dict]]],
                 out_path: Path, log_lines: list[str], log=print) -> Path:
    """
    `tab_sections` is [(heading, groups), ...] -- every selected section goes
    into ONE document. `groups` is one entry per report box:
        {"data_type": str,
         "items": [{"label": str, "observations": [str], "images": [Path]}]}
    An existing report for the same project is appended to rather than replaced.
    """
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.shared import Pt

    out_path.parent.mkdir(parents=True, exist_ok=True)

    if cfg.template and Path(cfg.template).exists():
        doc = _open_docx(Path(cfg.template))          # user-supplied styling
        fresh = False
    else:
        doc = Document()
        fresh = True
        # match the BT / Telepass page setup so charts are sized identically
        from docx.shared import Cm, Inches as _In
        s0 = doc.sections[0]
        s0.page_width, s0.page_height = Cm(21.0), Cm(29.7)      # A4
        s0.left_margin = s0.right_margin = _In(1)
        s0.top_margin = s0.bottom_margin = _In(1)

    sec0 = doc.sections[0]
    usable = sec0.page_width - sec0.left_margin - sec0.right_margin

    if fresh:
        title = doc.add_paragraph()
        run = title.add_run(
            f"{datetime.now():%Y} - {datetime.now():%B} - {cfg.project} - "
            f"Data Sanity Check: Action")
        run.bold = True
        run.font.name = BODY_FONT
        run.font.size = Pt(12)
        add_info_table(doc, {
            "Organization name:": cfg.org,
            "Brand name:": cfg.brand,
            "Model name:": cfg.project,
            "Model slug:": cfg.schema,
            "UIDs:": "",
            "Required action:": "Data Sanity and Format Checks",
            "Period impacted:": period_label(cfg),
            "Data availability:": "",
            "Notes:": "",
        })
        add_toc(doc)

    for heading, groups in tab_sections:
        doc.add_heading(heading, level=1)
        for n, grp in enumerate(groups, 1):
            variable = grp["data_type"]
            # Heading 2 puts the variable in the Contents index. Skip it when
            # the section has a single box named after the section itself.
            if not (len(groups) == 1 and variable.upper() == heading.upper()):
                doc.add_heading(f"{n}. {variable}", level=2)

            # one box per UID, immediately followed by that UID's own charts,
            # so the text and the images it describes sit together
            for i, item in enumerate(grp["items"], 1):
                label = item.get("label") or variable
                add_box(doc, variable, cfg.delivered_by,
                        [f"Variable: {variable}", QUESTION_LEAD,
                         f"{i}. UID: {label}", *item["observations"]], "")
                # Word fuses back-to-back tables into one; a UID whose images
                # all went missing would otherwise merge into the next box
                doc.add_paragraph()
                for k, img in enumerate(item.get("images", []), 1):
                    path = img.get("path") if isinstance(img, dict) else img
                    trace = img.get("trace", "") if isinstance(img, dict) else ""
                    if not path or not Path(path).exists():
                        continue
                    # don't repeat the series name when the label already has it
                    suffix = f"  -  {trace}" if trace and trace not in label else ""
                    cap = doc.add_paragraph()
                    cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = cap.add_run(f"{n}.{i}.{k}  {label}{suffix}")
                    run.bold = True
                    run.font.name = BODY_FONT
                    run.font.size = Pt(9)
                    pic = doc.add_paragraph()
                    pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    pic.add_run().add_picture(str(path), width=usable)

    mark_fields_dirty(doc)
    return save_docx_safely(doc, out_path, log)
# --------------------------------------------------------------------------- #
# The automation itself (everything AFTER the user clicks Load Data)
# --------------------------------------------------------------------------- #

def list_feed_options(bot: Bot, label: str) -> tuple[list[dict], str]:
    """Every option of the Feed Name selector, plus whatever is selected now."""
    try:
        res = bot.d.execute_script(JS_LIST_OPTIONS, label) or {}
    except JavascriptException as e:
        bot.log(f"WARN: could not read the '{label}' selector ({e})")
        return [], ""
    if not res.get("found"):
        bot.log(f"WARN: no '{label}' selector found")
        return [], ""
    opts = res.get("options") or []
    bot.log(f"'{res.get('label')}' ({res.get('widget')}): {len(opts)} option(s) "
            f"- {', '.join(o['text'] for o in opts[:12])}"
            + (" ..." if len(opts) > 12 else ""))
    return opts, res.get("current") or ""


def select_feed(bot: Bot, label: str, value: str) -> bool:
    try:
        res = bot.d.execute_script(JS_SELECT_ONE, value, label) or {}
    except JavascriptException as e:
        bot.log(f"  WARN: could not select '{value}' ({e})")
        return False
    if not res.get("ok"):
        bot.log(f"  WARN: '{label}' selector vanished while selecting '{value}'")
        return False
    return True


def select_filter_values(bot: Bot, wanted: tuple[str, ...],
                         label: str = "filter values") -> dict:
    """Pick specific options - Media wants 'net + impressions'."""
    try:
        res = bot.d.execute_script(JS_SELECT_VALUES, list(wanted), label) or {}
    except JavascriptException as e:
        bot.log(f"WARN: could not read the '{label}' selector ({e})")
        return {"mode": "error"}
    if res.get("mode") == "not-found":
        bot.log(f"WARN: no '{label}' selector found - leaving it as it is")
        return res
    bot.log(f"'{res.get('label') or label}' options: {res.get('available')}")
    if not res.get("matched"):
        bot.log(f"WARN: none of {list(wanted)} matched - the current selection "
                f"{res.get('before')} is left alone")
    else:
        bot.log(f"  selected {res.get('after')}")
    return res


def read_context(bot: Bot) -> dict:
    try:
        return bot.d.execute_script(JS_READ_CONTEXT) or {}
    except JavascriptException:
        return {}


def group_sections(sections: list[dict], profile: TabProfile,
                   bot: Bot) -> list[dict]:
    """
    Turn per-chart results into report variables, each holding its UIDs.
    Every trace of one chart (count + value, impressions + net) lands in the
    same UID, so the report shows that UID's box and then its own images.

      single      -> one variable for the whole tab, UIDs inside (Objective)
      by_variable -> one box per variable, UIDs inside    (Non-media)
                     'dispwofeatincrcaseper_bulleitwhiskey'
                       -> box 'dispwofeatincrcaseper', UID 'bulleitwhiskey'
      by_channel  -> one box per paid* channel, UIDs inside   (Media)
                     'flavors_googleads_paidsearch_tra'
                       -> box 'paidsearch', UID the full chart name
    """
    per_chart: dict[str, int] = {}
    for s in sections:
        per_chart[s["chart"]] = per_chart.get(s["chart"], 0) + 1

    order: list[str] = []
    buckets: dict[str, dict[str, dict]] = {}
    for s in sections:
        if profile.group_mode == "single":
            box = profile.data_type or profile.heading
        elif profile.group_mode == "by_channel":
            box = channel_of(s["chart"])          # paidsearch, paidsocial, ...
        elif profile.group_mode == "by_feed":
            box = s.get("feed") or split_title(s["chart"])[0]
        else:
            box = split_title(s["chart"])[0]
        uid = s["chart"]                          # the UID is the chart title
        if box not in buckets:
            buckets[box] = {}
            order.append(box)
        entry = buckets[box].setdefault(
            uid, {"label": uid, "observations": [], "images": []})
        # every trace of one chart shares a UID entry, so its images stay together
        if per_chart[s["chart"]] > 1 and s["trace"]:
            entry["observations"].append(f"{s['trace']}:")
        entry["observations"].extend(s["observations"])
        entry["images"].append({"path": s["image"], "trace": s["trace"]})

    bot.log(f"grouped {len(sections)} image(s) into {len(order)} box(es): "
            + ", ".join(f"{b} ({len(buckets[b])} UID)" for b in order))
    return [{"data_type": b, "items": list(buckets[b].values())} for b in order]


class Disconnected(Exception):
    """Shiny dropped the session and it could not be rebuilt."""


def snapshot_session(bot: Bot, cfg: Config) -> None:
    """
    Remember what the user typed, so a dropped session can be rebuilt without
    them. The token is only ever held in memory for this run.
    """
    try:
        tok = bot.d.execute_script(JS_READ_TOKEN) or ""
    except WebDriverException:
        tok = ""
    cfg.token_cache = tok
    bot.log("session snapshot: token "
            + (f"captured ({len(tok)} chars)" if tok else "NOT readable")
            + f", org '{cfg.org}', schema '{cfg.schema}'")
    if not tok:
        bot.log("  note: without the token a dropped session cannot be rebuilt "
                "automatically - the run will pause and wait for you instead")


def _wait_for_login(bot: Bot, cfg: Config) -> bool:
    """If the reload lands on Keycloak, wait for the human to sign in."""
    drv = bot.d
    if not is_login_page(drv.current_url or ""):
        return True
    bot.log("  the reload landed on the sign-in page - waiting for you to "
            "authenticate (the run resumes on its own afterwards)")
    deadline = time.time() + cfg.reconnect_wait
    told = 0
    while time.time() < deadline:
        bot.check_cancel()
        if not is_login_page(drv.current_url or ""):
            bot.log("  signed in")
            bot.wait_idle(timeout=cfg.timeout)
            return True
        left = int(deadline - time.time())
        if left // 60 != told:
            told = left // 60
            bot.log(f"  waiting for sign-in... {told} minute(s) left")
        time.sleep(2)
    return False


def _rebuild_session(bot: Bot, cfg: Config) -> bool:
    """
    Press the app's own Reload link, wait out any sign-in, then put the token,
    project, model and Load Data back exactly as they were.
    """
    drv = bot.d
    try:
        how = drv.execute_script(JS_CLICK_RELOAD)
        bot.log(f"  pressed the app's Reload ({how})")
        if how == "none":
            drv.get(cfg.url)
            bot.log("  no Reload link found - reloaded the page instead")
    except WebDriverException as exc:
        bot.log(f"  reload failed: {exc}")
        return False

    time.sleep(2.0)
    bot.wait_idle(timeout=cfg.timeout)
    if not _wait_for_login(bot, cfg):
        return False

    tok_el = bot.find("api_token", timeout=45, required=False)
    if tok_el is None:
        bot.log("  no API token field after the reload")
        return False
    if not cfg.token_cache:
        bot.log("  no token remembered - cannot rebuild automatically")
        return False
    try:
        tok_el.clear()
        tok_el.send_keys(cfg.token_cache)
    except WebDriverException as exc:
        bot.log(f"  could not type the token: {exc}")
        return False
    bot.log("  token re-entered")

    sub = bot.find("submit_token", timeout=20, required=False)
    if sub is not None:
        bot.click(sub, "Submit")
        bot.wait_idle()

    # put the project and model back, by option value
    for label, value, name in (("organization", cfg.org_value or cfg.org, "project"),
                               ("schema", cfg.schema_value or cfg.schema, "model")):
        if not value:
            continue
        if select_feed(bot, label, value):
            bot.log(f"  {name} restored: {value}")
        else:
            bot.log(f"  WARN: could not restore the {name} ({value})")
        bot.wait_idle()

    load = bot.find("load_data", timeout=30, required=False)
    if load is None:
        bot.log("  no Load Data button after the reload")
        return False
    bot.click(load, "Load Data")
    bot.wait_idle(quiet=2.0, timeout=cfg.plot_timeout)

    deadline = time.time() + 180
    while time.time() < deadline:
        bot.check_cancel()
        if len(bot.visible_tabs()) >= 4 and not bot.disconnected():
            bot.log("  data reloaded - session is usable again")
            return True
        time.sleep(1.5)
    bot.log(f"  rebuilt but still not usable (tabs={bot.visible_tabs()[:6]})")
    return False


def recover_session(bot: Bot, cfg: Config, tab: str) -> bool:
    """
    Deal with a Shiny disconnect: try to rebuild the session ourselves, and if
    that is not possible, wait for the user to do it. Returns True once the app
    is usable again and the wanted tab is open.
    """
    bot.log("")
    bot.log("!! Shiny disconnected from the server "
            "(this is the app, not the automation)")

    if cfg.token_cache:
        for attempt in (1, 2):
            bot.check_cancel()
            bot.log(f"reconnect attempt {attempt}/2 - reloading and signing back in")
            if _rebuild_session(bot, cfg):
                bot.log("reconnected automatically")
                return _reopen_tab(bot, tab)
            time.sleep(3)

    # nothing else for it: wait for a human
    span = (f"{cfg.reconnect_wait // 60} minutes" if cfg.reconnect_wait >= 60
            else f"{cfg.reconnect_wait} seconds")
    bot.log(f"waiting up to {span} for you to press "
            f"Reload in the browser and load the platform again "
            f"(token, org, schema, Load Data). The run continues by itself once "
            f"the app responds.")
    deadline = time.time() + cfg.reconnect_wait
    told = 0
    while time.time() < deadline:
        bot.check_cancel()
        if not bot.disconnected() and bot.visible_tabs():
            bot.log("the app is back - carrying on")
            return _reopen_tab(bot, tab)
        left = int(deadline - time.time())
        if left // 60 != told:
            told = left // 60
            bot.log(f"  still waiting... {told} minute(s) left")
        time.sleep(2)
    bot.log("gave up waiting for the app to come back")
    return False


def _reopen_tab(bot: Bot, tab: str) -> bool:
    key = {"Summary": "tab_summary", "Objective": "tab_objective",
           "Non-media": "tab_nonmedia", "Media": "tab_media"}.get(tab)
    if not key:
        return True
    el = bot.find(key, timeout=30, required=False)
    if el is None:
        bot.log(f"  could not reopen the '{tab}' tab after reconnecting")
        return False
    bot.click(el, f"{tab} tab")
    bot.wait_idle()
    return True


def ensure_connected(bot: Bot, cfg: Config, tab: str) -> None:
    """Raise Disconnected if the session is gone and cannot be restored."""
    if not bot.disconnected():
        return
    if not recover_session(bot, cfg, tab):
        raise Disconnected(
            "Shiny disconnected and the session could not be restored. "
            "Everything captured before the drop has been kept.")


def record_dates(cfg: Config, dfrom: str, dto: str, tag: str = "") -> None:
    """Remember a range we actually saw on the page, for the report header."""
    if not (dfrom or dto):
        return
    seen = getattr(cfg, "dates_seen", None)
    if seen is None:
        return
    entry = (dfrom, dto, tag)
    if entry not in seen:
        seen.append(entry)


def period_label(cfg: Config) -> str:
    """'Period impacted' = the span the app actually plotted, across all tabs."""
    seen = list(getattr(cfg, "dates_seen", None) or [])
    froms = sorted(f for f, _t, _g in seen if f)
    tos = sorted(t for _f, t, _g in seen if t)
    start = froms[0] if froms else (cfg.date_from or "").strip()
    end = tos[-1] if tos else (cfg.date_to or "").strip()
    if start and end:
        return f"{start} to {end}"
    return start or end or ""


def apply_date_range(bot: Bot, cfg: Config, tag: str = "") -> tuple[str, str]:
    """
    Take the Date Range from the page, not from a constant.

    The app fills its own Date Range from the data it loaded, and that range
    differs per model and per feed, so it is read every cycle. The From/To
    boxes in the run bar are an override: blank means "use what the app has",
    and anything typed is pushed in and then read back to prove it stuck.
    """
    live = bot.read_date_range()
    if not live:
        bot.log("no Date Range control on this tab - nothing to read or set")
        record_dates(cfg, (cfg.date_from or "").strip(),
                     (cfg.date_to or "").strip(), tag)
        return ((cfg.date_from or "").strip(), (cfg.date_to or "").strip())

    where = "the active tab" if live.get("active") else "the page"
    bounds = ""
    if live.get("min") or live.get("max"):
        bounds = (f"  [app allows {live.get('min') or '?'} .. "
                  f"{live.get('max') or '?'}]")
    bot.log(f"date range read off {where}: {live.get('from_iso') or '?'} to "
            f"{live.get('to_iso') or '?'}{bounds}")

    want_from = (cfg.date_from or "").strip()
    want_to = (cfg.date_to or "").strip()
    if not want_from and not want_to:
        eff = (live.get("from_iso", ""), live.get("to_iso", ""))
        bot.log("no override in the run bar - keeping the app's own range")
    else:
        if want_from:
            bot.set_date("date_from", want_from)
        if want_to:
            bot.set_date("date_to", want_to)
        bot.wait_idle()
        after = bot.read_date_range()
        eff = (after.get("from_iso") or live.get("from_iso", ""),
               after.get("to_iso") or live.get("to_iso", ""))
        bot.log(f"date range in use: {eff[0] or '?'} to {eff[1] or '?'}")
    record_dates(cfg, eff[0], eff[1], tag)
    return eff


def _plot_and_export(bot: Bot, cfg: Config, profile: TabProfile, images: Path,
                     tag: str) -> list[dict]:
    """One full plot cycle: options -> Go Plot -> isolate + export every trace."""
    drv = bot.d
    sections: list[dict] = []

    # date range: read off the page every cycle, overridden only if asked
    apply_date_range(bot, cfg, tag)
    bot.wait_idle()

    # Plots -> Time-Series Plots
    pl = bot.find("tab_plots", timeout=20, required=False)
    if pl:
        bot.click(pl, "Plots sub-tab")
        bot.wait_idle()
    ts = bot.find("tab_timeseries", timeout=10, required=False)
    if ts:
        bot.click(ts, "Time-Series Plots")
        bot.wait_idle()

    # plot options, straight from the profile -- verified, because a control
    # that silently fails to change produces a chart of the wrong type
    bot.step(f"plot options for '{tag}'")
    ok = [bot.set_marker(profile.marker)]
    for label, want in profile.checkboxes:
        ok.append(bot.set_checkbox_labelled(label, want))
    if not all(ok):
        bot.log("WARN: not every plot option could be set - the chart may not "
                "match the requested marker/checkbox settings")
    bot.wait_idle()

    bot.check_cancel()
    bot.step("click Go Plot and wait for the charts")
    bot.click("go_plot")
    try:
        bot.wait_plots(timeout=profile.plot_timeout, settle=profile.settle)
    except TimeoutException:
        if bot.disconnected():
            raise Disconnected("the session dropped while waiting for charts")
        bot.log(f"WARN: no charts appeared for '{tag}' - treating it as no data")
        return []

    info = drv.execute_script(JS_TRACE_INFO)
    bot.progress(kind="charts", tag=tag, chart_n=len(info))
    bot.log(f"{len(info)} chart(s); first traces: " + json.dumps(
        [[t["name"] for t in c["traces"]] for c in info[:4]]))

    hidden = {norm(h) for h in profile.hide}
    skipped: list[str] = []
    total = len(info)
    seq = 0
    safe_tag = re.sub(r"[^A-Za-z0-9_-]+", "_", tag).strip("_")[:28]
    drv.set_script_timeout(90)            # once per cycle, not per download
    started = time.time()

    for chart in info:
        bot.check_cancel()
        gi = chart["index"]
        title = chart["title"] or f"chart_{gi + 1}"
        names = {norm(t["name"]): t["name"] for t in chart["traces"]}
        wanted = (list(profile.export) if profile.export
                  else [t["name"] for t in chart["traces"]
                        if norm(t["name"]) not in hidden])
        bot.step(f"[{tag}] chart {gi + 1}/{total}: '{title}' -> {wanted}")

        try:
            raw = drv.execute_script(JS_EXTRACT_DATA, gi)
        except WebDriverException as exc:
            skipped.append(title)
            bot.log(f"  WARN: could not read chart {gi} ({exc}) - skipped")
            continue

        for want in wanted:
            bot.check_cancel()
            wn = norm(want)
            if wn not in names:
                bot.log(f"  no '{want}' trace here - skipped")
                continue
            # charts get re-rendered; make sure the index still points here
            gi = resolve_plot_index(bot, title, gi)
            if not set_visible_traces(bot, gi, [names[wn]]):
                bot.log(f"  skipping the '{want}' image for '{title}' rather "
                        f"than saving a blank chart")
                skipped.append(f"{title}/{want}")
                continue

            seq += 1
            # unique, self-describing name: nothing can overwrite anything else
            stem = re.sub(r"[^A-Za-z0-9_-]+", "_",
                          f"{title}_{want}").strip("_")[:80]
            slug = f"{seq:03d}_{safe_tag}_{stem}" if safe_tag else f"{seq:03d}_{stem}"
            img = download_plot_png(bot, gi, slug, images)
            bot.log(f"    image: {img.name}")
            bot.progress(kind="image", tag=tag, chart=title, trace=want,
                         chart_i=gi + 1, chart_n=total, images=seq)

            tr = next((t for t in raw["traces"] if norm(t["name"]) == wn), None)
            if tr:
                try:
                    prose = narrative_observations(tr["x"], tr["y"], names[wn])
                    stats = analyze_series(tr["x"], tr["y"], names[wn])
                except Exception as exc:          # never lose a whole run
                    bot.log(f"  WARN: could not summarise '{names[wn]}' "
                            f"({type(exc).__name__}: {exc}) - chart still kept")
                    prose = [f"Automatic summary unavailable "
                             f"({type(exc).__name__}). Please review the chart."]
                    stats = prose
            else:
                prose = ["No data extracted for this trace."]
                stats = prose
            for line in prose:
                bot.log(f"    - {line}")
            sections.append({"chart": title, "trace": names[wn], "feed": tag,
                             "image": str(img), "observations": prose,
                             "stats": stats})

        # Deliberately not restoring both series here: it is purely cosmetic and
        # costs a restyle plus a settle on every chart, which is minutes across
        # a hundred charts.

    took = time.time() - started
    bot.log(f"[{tag}] {seq} image(s) in {took:.0f}s"
            + (f" ({took / seq:.1f}s each)" if seq else ""))
    if skipped:
        bot.log(f"{len(skipped)} image(s) skipped: {', '.join(skipped[:10])}"
                + (" ..." if len(skipped) > 10 else ""))
    return sections


def capture_tab(bot: Bot, cfg: Config, tab: str, base: Path,
                run_id: str) -> list[dict]:
    """Drive one Data App tab and return its report boxes."""
    profile = TAB_PROFILES[tab]
    images = base / "charts" / f"{tab.lower().replace(' ', '_')}_{run_id}"
    images.mkdir(parents=True, exist_ok=True)
    set_download_dir(bot.d, images)

    bot.log("")
    bot.log(f"########## {tab} -> {profile.heading} ##########")
    bot.progress(kind="section", tag=tab)
    bot.log(f"charts -> {images}")

    bot.step(f"open the '{tab}' tab")
    key = {"Summary": "tab_summary", "Objective": "tab_objective",
           "Non-media": "tab_nonmedia", "Media": "tab_media"}[tab]
    el = bot.find(key, timeout=25, required=False)
    if el:
        bot.click(el, f"{tab} tab")
        bot.wait_idle()
    else:
        bot.log(f"WARN: '{tab}' tab not found. Tabs: {bot.visible_tabs()}")

    if cfg.debug_dom:
        bot.dump_dom(base / f"dom_{tab.lower()}_{run_id}.json")

    sections: list[dict] = []

    try:
        _capture_body(bot, cfg, profile, tab, images, sections)
    except Disconnected as exc:
        bot.log(f"!! {exc}")
        bot.log(f"keeping the {len(sections)} image(s) captured before the drop")
    except Cancelled:
        bot.was_cancelled = True
        bot.log(f"cancelled - keeping the {len(sections)} image(s) from {tab}")

    (base / f"observations_{tab.lower()}_{run_id}.json").write_text(
        json.dumps(sections, indent=2), encoding="utf-8")
    if not sections:
        bot.log(f"WARN: {tab} produced no chart images")
        return []
    return group_sections(sections, profile, bot)


def _capture_body(bot: Bot, cfg: Config, profile: TabProfile, tab: str,
                  images: Path, sections: list[dict]) -> None:
    if profile.iterate_feeds:
        # Feed Name is a SINGLE select: capture one feed, then the next, until
        # every option in the dropdown has been done.
        opts, current = list_feed_options(bot, profile.feed_label)
        if not opts:
            bot.log("no feed list found - running once with the current selection")
            sections.extend(_plot_and_export(bot, cfg, profile, images,
                                             current or tab))
        else:
            # Resume-capable loop: a disconnect rebuilds the session and retries
            # the SAME feed, then carries on with the ones still to come.
            i, tries = 0, 0
            empty: list[str] = []
            while i < len(opts):
                bot.check_cancel()
                opt = opts[i]
                name = opt.get("text") or opt.get("value")
                bot.log("")
                bot.step(f"===== feed {i + 1}/{len(opts)}: {name} =====")
                bot.progress(kind="feed", tag=name, feed_i=i + 1, feed_n=len(opts))
                try:
                    ensure_connected(bot, cfg, tab)
                    if not select_feed(bot, profile.feed_label, opt["value"]):
                        i += 1
                        tries = 0
                        continue
                    bot.wait_idle()
                    n_ins = bot.insertions()
                    if n_ins == 0:
                        # nothing to plot for this feed: Go Plot would just spin
                        bot.log(f"feed '{name}': 0 insertions - no data, skipped")
                        empty.append(name)
                        i += 1
                        tries = 0
                        continue
                    if n_ins > 0:
                        bot.log(f"feed '{name}': {n_ins} insertion(s)")
                    got = _plot_and_export(bot, cfg, profile, images, name)
                    sections.extend(got)
                    bot.log(f"feed '{name}': {len(got)} image(s)")
                    i += 1
                    tries = 0
                except Cancelled:
                    raise
                except Disconnected:
                    tries += 1
                    # drop anything half-captured for this feed before retrying
                    sections[:] = [x for x in sections if x.get("feed") != name]
                    if tries > 3:
                        bot.log(f"!! '{name}' keeps dropping - giving up on the "
                                f"rest of {tab}")
                        raise
                    bot.log(f"resuming at feed {i + 1}/{len(opts)} "
                            f"('{name}') after the disconnect")
                except Exception as exc:
                    if bot.disconnected() and tries < 3:
                        tries += 1
                        sections[:] = [x for x in sections
                                       if x.get("feed") != name]
                        continue
                    bot.log(f"!! feed '{name}' failed "
                            f"({type(exc).__name__}: {exc}) - skipped")
                    empty.append(f"{name} ({type(exc).__name__})")
                    i += 1
                    tries = 0
            if empty:
                bot.log(f"{len(empty)} feed(s) skipped: {', '.join(empty)}")
                bot.skipped_feeds.extend(empty)

    else:
        for attempt in (1, 2, 3):
            try:
                ensure_connected(bot, cfg, tab)
                if profile.filter_values:
                    bot.step(f"Filter Values -> {list(profile.filter_values)}")
                    select_filter_values(bot, profile.filter_values)
                    bot.wait_idle()
                sections.extend(_plot_and_export(bot, cfg, profile, images, tab))
                break
            except (Cancelled, Disconnected):
                if attempt == 3:
                    raise
                sections.clear()
                bot.log(f"restarting {tab} after the disconnect "
                        f"(attempt {attempt + 1}/3)")


def run_after_load(drv, cfg: Config) -> Path:
    """Capture every selected section, then write them into ONE document."""
    bot = Bot(drv, cfg)

    ctx = read_context(bot)
    cfg.org = ctx.get("organization") or cfg.org
    cfg.schema = ctx.get("schema") or cfg.schema
    if not cfg.project:
        cfg.project = cfg.schema or cfg.org or "project"
    bot.log(f"organization: {cfg.org or '(unknown)'}")
    bot.log(f"schema      : {cfg.schema or '(unknown)'}")
    bot.log(f"project     : {cfg.project}")

    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    # one folder per run, so re-running the same model never overwrites the
    # previous set of charts, JSON or report
    root = Path(cfg.outdir) if cfg.outdir else DOWNLOADS_ROOT / cfg.project
    base = root / f"run_{run_id}"
    base.mkdir(parents=True, exist_ok=True)
    cfg.run_dir = base
    bot.set_log_file(base / f"run_log_{run_id}.txt")
    snapshot_session(bot, cfg)      # so a dropped session can be rebuilt

    wanted = [t for t in SECTION_ORDER if t in (cfg.sections or ())]
    if not wanted:
        raise RuntimeError("No sections selected - tick at least one of "
                           "Objective, Non-media or Media.")
    bot.log(f"sections to capture: {', '.join(wanted)}")

    payload: list[tuple[str, list[dict]]] = []
    failures: list[str] = []
    try:
        for tab in wanted:
            try:
                groups = capture_tab(bot, cfg, tab, base, run_id)
                if groups:
                    payload.append((TAB_PROFILES[tab].heading, groups))
                if bot.was_cancelled:
                    bot.log("stopping here - writing what has been captured")
                    break
            except Cancelled:
                raise
            except Exception as exc:
                # one bad section must not lose the sections that worked
                failures.append(f"{tab}: {type(exc).__name__}: {exc}")
                bot.log(f"!! {tab} FAILED: {type(exc).__name__}: {exc}")
                try:
                    drv.save_screenshot(str(base / f"FAILURE_{tab}_{run_id}.png"))
                    (base / f"FAILURE_{tab}_{run_id}.html").write_text(
                        drv.page_source, encoding="utf-8")
                    bot.dump_dom(base / f"FAILURE_{tab}_dom_{run_id}.json")
                except Exception:
                    pass
    except Cancelled:
        if not payload:
            raise
        bot.log("cancelled - writing the sections captured so far")

    if not payload:
        raise RuntimeError(
            "Nothing was captured.\n" + ("\n".join(failures) if failures else ""))

    out = base / (f"{cfg.project} Data Sanity Check Report_{run_id}.docx")
    doc = build_report(cfg, payload, out, bot.log_lines, log=bot.log)
    if bot.skipped_feeds:
        bot.log("feeds with no data (skipped): " + ", ".join(bot.skipped_feeds))
    bot.log("")
    bot.log(f"report written: {doc}")
    if failures:
        bot.log("sections that failed (the rest are in the report): "
                + "; ".join(failures))
    return doc

# --------------------------------------------------------------------------- #
# UI - a floating "dynamic island": borderless, always on top, draggable,
# resizable, sitting inside a live RGB gradient rim.
#
# It walks the setup one step at a time, each step replacing the last:
#     Start  ->  Auth done  ->  Project  ->  Model  ->  (Load Data)  ->  Run bar
# The run bar carries the dates, the section chips, Run/Cancel and live progress.
# --------------------------------------------------------------------------- #

import tkinter as tk

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

INK = "#000000"          # island body: true black, like the Dynamic Island
INK2 = "#1B1B24"         # raised surface
INK3 = "#262633"         # hover
TEXT = "#F5F5F7"
MUTED = "#9A9AA6"
ACCENT = "#0A84FF"
GOOD = "#32D74B"
WARN = "#FFD60A"
BAD = "#FF453A"


FONT = "Segoe UI"
RADIUS = 60          # near-stadium ends; the log-open slab stays soft
LOG_H = 190          # extra logical height when the log panel is open
KEY = "#010203"      # keyed out to transparent, so the corners are really round
PAD = 2              # breathing room so the rounded corners are not clipped
INSET = 11           # widgets sit this far inside the pill, clear of the curve
BORDER = "#1C1C24"   # a hairline edge instead of a shadow: no repaint, no
                     # screen capture, nothing to go stale behind the window
MORPH_MS = 220       # island resize animation
MORPH_STEPS = 14


class DatePopover(ctk.CTkToplevel):
    """
    A month calendar dropped under a date box. Picking a day writes it into
    the box as YYYY-MM-DD; Clear empties it, which means "use the app's own
    range". The box itself stays typeable.
    """

    def __init__(self, master, entry: ctk.CTkEntry, fallback: str,
                 on_close) -> None:
        super().__init__(master)
        self.entry = entry
        self.on_close = on_close
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(fg_color=KEY)
        try:
            self.attributes("-transparentcolor", KEY)
        except tk.TclError:    # no colour key on this platform: square corners
            self.configure(fg_color=INK)
        self.bind("<Escape>", lambda e: self.close())

        today = datetime.now().date()
        self.today = today
        self.picked = self._parse(entry.get()) or None
        start = self.picked or self._parse(fallback) or today
        self.year, self.month = start.year, start.month

        frame = ctk.CTkFrame(self, fg_color=INK, corner_radius=18,
                             border_width=1, border_color=BORDER)
        frame.pack(fill="both", expand=True)
        head = ctk.CTkFrame(frame, fg_color="transparent")
        head.pack(fill="x", padx=10, pady=(10, 4))
        nav = dict(width=28, height=24, corner_radius=12, fg_color=INK2,
                   hover_color=INK3, text_color=TEXT,
                   font=ctk.CTkFont(FONT, 12, "bold"))
        ctk.CTkButton(head, text="\u2039", command=lambda: self._shift(-1),
                      **nav).pack(side="left")
        ctk.CTkButton(head, text="\u203a", command=lambda: self._shift(1),
                      **nav).pack(side="right")
        self.title_lbl = ctk.CTkLabel(head, text="", text_color=TEXT,
                                      font=ctk.CTkFont(FONT, 12, "bold"))
        self.title_lbl.pack(side="left", expand=True)

        grid = ctk.CTkFrame(frame, fg_color="transparent")
        grid.pack(padx=10)
        for c, name in enumerate(("Mo", "Tu", "We", "Th", "Fr", "Sa", "Su")):
            ctk.CTkLabel(grid, text=name, width=32, height=18,
                         text_color=MUTED,
                         font=ctk.CTkFont(FONT, 10)).grid(row=0, column=c)
        self.cells: list[ctk.CTkButton] = []
        for i in range(42):
            b = ctk.CTkButton(grid, text="", width=32, height=26,
                              corner_radius=13, border_width=0,
                              font=ctk.CTkFont(FONT, 11))
            b.grid(row=1 + i // 7, column=i % 7, padx=1, pady=1)
            self.cells.append(b)

        foot = ctk.CTkFrame(frame, fg_color="transparent")
        foot.pack(fill="x", padx=10, pady=(6, 10))
        small = dict(width=70, height=24, corner_radius=12, fg_color=INK2,
                     hover_color=INK3, text_color=TEXT,
                     font=ctk.CTkFont(FONT, 11))
        ctk.CTkButton(foot, text="Clear", command=lambda: self._set(""),
                      **small).pack(side="left")
        ctk.CTkButton(foot, text="Today",
                      command=lambda: self._set(today.isoformat()),
                      **small).pack(side="right")

        self._render()
        self.update_idletasks()
        self.geometry(f"+{entry.winfo_rootx()}"
                      f"+{entry.winfo_rooty() + entry.winfo_height() + 6}")

    @staticmethod
    def _parse(text: str):
        try:
            return datetime.strptime(text.strip(), "%Y-%m-%d").date()
        except ValueError:
            return None

    def _shift(self, months: int) -> None:
        m = self.month - 1 + months
        self.year, self.month = self.year + m // 12, m % 12 + 1
        self._render()

    def _render(self) -> None:
        import calendar
        self.title_lbl.configure(text=f"{calendar.month_name[self.month]} "
                                      f"{self.year}")
        days = [d for week in calendar.Calendar().monthdatescalendar(
            self.year, self.month) for d in week]
        days += [None] * (42 - len(days))
        for b, d in zip(self.cells, days):
            if d is None or d.month != self.month:
                b.configure(text="", state="disabled", fg_color="transparent",
                            border_width=0)
                continue
            sel = d == self.picked
            b.configure(text=str(d.day), state="normal",
                        fg_color=ACCENT if sel else "transparent",
                        hover_color="#0A6FD8" if sel else INK3,
                        text_color="#FFFFFF" if sel else TEXT,
                        border_width=1 if d == self.today and not sel else 0,
                        border_color=ACCENT,
                        command=lambda d=d: self._set(d.isoformat()))

    def _set(self, iso: str) -> None:
        self.entry.delete(0, "end")
        if iso:
            self.entry.insert(0, iso)
        self.close()

    def close(self) -> None:
        self.on_close()
        self.destroy()


def _hex(rgb) -> str:
    return "#%02x%02x%02x" % tuple(int(c) for c in rgb)


def _rgb(hexcolor: str) -> tuple[int, int, int]:
    h = hexcolor.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _dist(a, b) -> int:
    return sum(abs(int(x) - int(y)) for x, y in zip(a, b))


class Island(ctk.CTk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Data Sanity")
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self._bg_key = None
        self._dragging = False
        self._user_placed = False
        self._morph_job = None
        self.configure(fg_color=KEY)
        try:
            self.attributes("-transparentcolor", KEY)
            self._keyed = True
        except Exception:      # no key colour available: plain dark window
            self._keyed = False
            self.configure(fg_color=INK)

        self.q: queue.Queue[tuple[str, object]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.cancel_event = threading.Event()
        self.driver = None
        self.bot: Bot | None = None
        self.cfg = Config()
        # the range the Data App itself reports, shown as the placeholder so an
        # empty box keeps meaning "whatever the app has"
        self.app_dates: tuple[str, str] = ("", "")
        self.last_output: Path | None = None
        self.date_pop: DatePopover | None = None
        self.log_open = False
        self._drag = (0, 0)
        self._bg_img = None
        self._stage_h = 176

        w, h = 780, 96
        sc0 = self._scale()
        w = min(w, max(520, int(self.winfo_screenwidth() / sc0) - 24))
        self._wl_cur = w
        x0 = max(0, (self.winfo_screenwidth() - int(w * sc0)) // 2)
        self.geometry(f"{w}x{h}+{x0}+40")
        self.minsize(380, 96)

        # backdrop + shadow, painted as one image
        self.bg = tk.Label(self, bd=0, highlightthickness=0, bg=KEY)
        self.bg.place(x=0, y=0, relwidth=1, relheight=1)

        # the island body sits inside the rim. CTk refuses width/height in
        # place(), so a plain tk.Frame does the inset and the card fills it.
        edge = PAD + INSET
        self.card = tk.Frame(self, bg=INK, bd=0, highlightthickness=0)
        self.card.place(x=edge, y=edge, relwidth=1, relheight=1,
                        width=-2 * edge, height=-2 * edge)

        self._build_header()
        self.status = ctk.CTkLabel(self.card, text="", anchor="w",
                                   text_color=MUTED,
                                   font=ctk.CTkFont(FONT, 11))
        self.status.pack(side="bottom", fill="x", padx=29, pady=(0, 7))
        self.logbox = ctk.CTkTextbox(self.card, height=180, fg_color=INK2,
                                     border_width=0, text_color=MUTED,
                                     corner_radius=12, wrap="word",
                                     font=ctk.CTkFont("Consolas", 11))
        self.body = ctk.CTkFrame(self.card, fg_color="transparent")
        self.body.pack(side="top", fill="both", expand=True, padx=26, pady=(1, 2))

        self._grip()
        self._bind_drag(self.bg)
        self._bind_drag(self.card)
        self.bind("<Configure>", self._on_configure)
        self.bind("<Escape>", lambda e: self._quit())

        self.stage_start()
        self.after(60, self._first_paint)
        self.after(120, self._drain)

    def _first_paint(self):
        self._paint_bg()

    # ------------------------------------------------------------------ #
    # gradient backdrop
    # ------------------------------------------------------------------ #
    def _raise_grip(self):
        try:
            self.grip.lift()
            self.gripdot.lift()
        except Exception:
            pass

    def _paint_bg(self, recapture=False):
        """
        The island is a black capsule: round ends while it is compact, a
        softly rounded slab once it grows. No screen capture or DWM blur --
        the first went stale, and DWM adds Tk's colours onto the blur, which
        washed every colour out.
        """
        from PIL import Image, ImageDraw, ImageTk
        w, h = self.winfo_width(), self.winfo_height()
        if w < 60 or h < 60:
            return
        if (w, h) == self._bg_key:
            return
        self._bg_key = (w, h)

        pill = [PAD, PAD, w - PAD - 1, h - PAD - 1]
        r = max(12, min(int(RADIUS * self._scale()), (pill[3] - pill[1]) // 2))

        base = Image.new("RGB", (w, h), _rgb(KEY) if self._keyed else _rgb(INK))
        ImageDraw.Draw(base).rounded_rectangle(pill, radius=r, fill=_rgb(INK),
                                               outline=_rgb(BORDER), width=1)

        self._bg_img = ImageTk.PhotoImage(base)
        self.bg.configure(image=self._bg_img)

    def _restage(self, _e=None):
        self._paint_bg()

    def _on_configure(self, _e=None):
        self._paint_bg()
        wrap = max(240, self.winfo_width() - 2 * (PAD + INSET) - 40)
        for lbl in getattr(self, "_wrapped", []):
            try:
                lbl.configure(wraplength=wrap)
            except Exception:
                pass

    # ------------------------------------------------------------------ #
    # header: drag handle, log toggle, close
    # ------------------------------------------------------------------ #
    def _build_header(self):
        bar = ctk.CTkFrame(self.card, fg_color="transparent", height=26)
        bar.pack(side="top", fill="x", padx=24, pady=(8, 0))
        bar.pack_propagate(False)

        self.dot = ctk.CTkLabel(bar, text="●", text_color=ACCENT, width=10,
                                font=ctk.CTkFont(FONT, 11))
        self.dot.pack(side="left", padx=(4, 7))
        self.title_lbl = ctk.CTkLabel(bar, text="Data Sanity", anchor="w",
                                      text_color=TEXT,
                                      font=ctk.CTkFont(FONT, 12, "bold"))
        self.title_lbl.pack(side="left")

        ctk.CTkButton(bar, text="✕", width=24, height=21, corner_radius=10,
                      fg_color=INK2, hover_color=BAD, text_color=MUTED,
                      font=ctk.CTkFont(FONT, 10), command=self._quit
                      ).pack(side="right", padx=(5, 1))
        self.btn_log = ctk.CTkButton(bar, text="Log", width=40, height=21,
                                     corner_radius=10, fg_color=INK2,
                                     hover_color=INK3, text_color=MUTED,
                                     font=ctk.CTkFont(FONT, 10),
                                     command=self._toggle_log)
        self.btn_log.pack(side="right")

        self._bind_drag(bar)

    _INTERACTIVE = None      # filled in on first use

    def _bind_drag(self, widget):
        """
        Make the island draggable by any dead space. CTk widgets are composites
        - the visible surface is an inner canvas - so binding the CTk object
        alone never sees the click; every non-interactive descendant is bound.
        """
        if Island._INTERACTIVE is None:
            Island._INTERACTIVE = (ctk.CTkButton, ctk.CTkEntry,
                                   ctk.CTkOptionMenu, ctk.CTkTextbox,
                                   ctk.CTkProgressBar, ctk.CTkSlider,
                                   ctk.CTkCheckBox, ctk.CTkSwitch)
        if isinstance(widget, Island._INTERACTIVE):
            return
        if widget not in (getattr(self, "grip", None),
                          getattr(self, "gripdot", None)):
            try:
                widget.bind("<Button-1>", self._drag_start, add="+")
                widget.bind("<B1-Motion>", self._drag_move, add="+")
                widget.bind("<ButtonRelease-1>", self._drag_end, add="+")
                widget.configure(cursor="fleur")
            except Exception:
                pass
        for child in widget.winfo_children():
            self._bind_drag(child)

    def _drag_start(self, e):
        self._close_calendar()          # it would be left floating behind
        self._dragging = True
        self._drag = (e.x_root - self.winfo_x(), e.y_root - self.winfo_y())

    def _drag_move(self, e):
        # free movement -- only enough of the island is kept on screen to grab
        # it again. Clamping it to the full width was the "wall".
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x = max(140 - w, min(e.x_root - self._drag[0], sw - 140))
        y = max(0, min(e.y_root - self._drag[1], sh - 60))
        self._user_placed = True
        self.geometry(f"+{x}+{y}")

    def _drag_end(self, _e=None):
        self._dragging = False

    def _grip(self):
        # a real tk widget (CTk composites swallow the click) and big enough to
        # actually grab
        self.grip = tk.Frame(self, bg=INK, width=18, height=18,
                             cursor="bottom_right_corner")
        self.grip.place(relx=1.0, rely=1.0, anchor="se",
                        x=-(PAD + 26), y=-(PAD + 12))   # inside the curve
        self.gripdot = tk.Label(self.grip, text="◢", bg=INK, fg=INK3, bd=0,
                                font=(FONT, 8), cursor="bottom_right_corner")
        self.gripdot.place(relx=1.0, rely=1.0, anchor="se")
        for w in (self.grip, self.gripdot):
            w.bind("<Button-1>", self._resize_start)
            w.bind("<B1-Motion>", self._resize)

    def _resize_start(self, e):
        sc = self._scale()
        self._rs = (e.x_root, e.y_root,
                    int(self.winfo_width() / sc), int(self.winfo_height() / sc))

    def _resize(self, e):
        sc = self._scale()
        x0, y0, w0, h0 = getattr(self, "_rs", (e.x_root, e.y_root, 820, 176))
        wl = max(520, int(w0 + (e.x_root - x0) / sc))
        hl = max(140, int(h0 + (e.y_root - y0) / sc))
        self._stage_h = hl - (LOG_H if self.log_open else 0)
        self._wl_cur = wl
        self._apply_geometry(wl, hl, keep_center=False, animate=False)
        self._paint_bg()

    def _toggle_log(self):
        self.log_open = not self.log_open
        if self.log_open:
            self.logbox.pack(side="bottom", fill="both", expand=True,
                             padx=18, pady=(0, 6))
        else:
            self.logbox.pack_forget()
        self._fit()

    # ------------------------------------------------------------------ #
    # sizing: never clip the buttons again
    # ------------------------------------------------------------------ #
    def _scale(self) -> float:
        try:
            from customtkinter.windows.widgets.scaling import ScalingTracker
            return ScalingTracker.get_window_scaling(self) or 1.0
        except Exception:
            return 1.0

    def _max_wl(self) -> int:
        """Widest the island may be, in logical units, so it fits the screen."""
        return max(520, int(self.winfo_screenwidth() / self._scale()) - 24)

    @property
    def _wl(self) -> int:
        """Intended width in logical units (what geometry() expects)."""
        return max(520, min(self._max_wl(), self._wl_cur))

    def _apply_geometry(self, wl: int, hl: int, keep_center: bool = True,
                        animate: bool = True):
        """
        Resize about the island's own centre. geometry() takes logical units,
        winfo_* returns physical, so the centre is worked out in physical and
        only the size is handed back logical. Stage changes morph to the new
        size the way the Dynamic Island does; a drag on the grip is direct.
        """
        sc = self._scale()
        wl = max(520, min(wl, self._max_wl()))    # never wider than the screen
        self._wl_cur = wl
        x, y = self.winfo_x(), self.winfo_y()
        old_w = self.winfo_width()
        new_w = int(wl * sc)
        sw = self.winfo_screenwidth()
        if old_w <= 1:                    # not on screen yet: centre it
            x = max(0, (sw - new_w) // 2)
        elif keep_center and new_w != old_w:
            x = int(x + old_w / 2 - new_w / 2)
        if not self._user_placed and (x < 0 or x + new_w > sw):
            x = max(0, (sw - new_w) // 2)   # only auto-place before you move it
        x = max(140 - new_w, min(x, sw - 140))
        y = max(0, y)
        if self._morph_job is not None:
            self.after_cancel(self._morph_job)
            self._morph_job = None
        if not animate or old_w <= 1:
            self.geometry(f"{wl}x{hl}+{x}+{y}")
            return
        self._morph((self.winfo_width() / sc, self.winfo_height() / sc,
                     self.winfo_x(), self.winfo_y()), (wl, hl, x, y), 1)

    def _morph(self, a, b, step: int):
        t = step / MORPH_STEPS
        e = 1 - (1 - t) ** 3                       # ease-out cubic
        w, h, x, y = (int(round(p + (q - p) * e)) for p, q in zip(a, b))
        self.geometry(f"{w}x{h}+{x}+{y}")
        if step < MORPH_STEPS:
            self._morph_job = self.after(MORPH_MS // MORPH_STEPS,
                                         self._morph, a, b, step + 1)
        else:
            self._morph_job = None

    def _fit(self, h=None):
        """
        Size the island to the stage. geometry() takes logical units while
        winfo_* reports physical pixels, so the measured requirement is scaled
        back down before use -- mixing the two is what clipped the buttons.
        """
        self.update_idletasks()
        sc = self._scale()
        need = int((self.card.winfo_reqheight() + 2 * (PAD + INSET)) / sc) + 4
        h = max(h or self._stage_h, need)
        self._stage_h = h
        total = h + (LOG_H if self.log_open else 0)
        self._apply_geometry(self._wl, total)
        self.update_idletasks()
        self._bind_drag(self.card)     # stage widgets are new; rebind the drag
        self._raise_grip()
        self._paint_bg()

    # ------------------------------------------------------------------ #
    # builders
    # ------------------------------------------------------------------ #
    def _clear(self):
        self._close_calendar()
        self._wrapped = []
        for w in self.body.winfo_children():
            w.destroy()

    def _hint(self, text, color=MUTED, size=11):
        lbl = ctk.CTkLabel(self.body, text=text, anchor="w", justify="left",
                           height=16, text_color=color,
                           wraplength=max(220, self.winfo_width()
                                           - 2 * (PAD + INSET) - 34),
                           font=ctk.CTkFont(FONT, size))
        self._wrapped = getattr(self, "_wrapped", []) + [lbl]
        return lbl

    def _pill(self, parent, text, cmd, primary=True, width=130, color=ACCENT):
        return ctk.CTkButton(
            parent, text=text, command=cmd, width=width, height=29,
            corner_radius=15,
            fg_color=color if primary else INK2,
            hover_color="#0A6FD8" if primary else INK3,
            text_color="#FFFFFF" if primary else TEXT,
            font=ctk.CTkFont(FONT, 12, "bold"))

    def say(self, text, color=MUTED):
        self.status.configure(text=text, text_color=color)
        self.dot.configure(text_color=color if color != MUTED else ACCENT)

    def log(self, line):
        self.logbox.insert("end", str(line).rstrip() + "\n")
        self.logbox.see("end")

    # ------------------------------------------------------------------ #
    # stage 1 - start
    # ------------------------------------------------------------------ #
    def stage_start(self):
        self._clear()
        self.title_lbl.configure(text="Data Sanity")
        self._hint("Opens Chrome and walks you through sign-in, "
                   "project and model.").pack(fill="x", pady=(2, 7))
        row = ctk.CTkFrame(self.body, fg_color="transparent")
        row.pack(fill="x", pady=(0, 2))
        self._pill(row, "Start", self.do_open_browser, width=150).pack(side="left")
        self.say("ready")
        self._fit(126)

    def do_open_browser(self):
        self._clear()
        self._hint("Launching Chrome...").pack(fill="x", pady=(2, 7))
        self.say("opening the browser", ACCENT)
        self._fit(112)

        def job():
            if self.driver is not None:
                try:
                    self.driver.quit()
                except WebDriverException:
                    pass
            cfg = self._base_config()
            drv = build_driver(cfg, DOWNLOADS_ROOT / "_staging")
            self.driver = drv
            self.bot = Bot(drv, cfg)
            drv.get(cfg.url)
            return "browser"

        self._run(job)

    # ------------------------------------------------------------------ #
    # stage 2 - auth
    # ------------------------------------------------------------------ #
    def stage_auth(self):
        self._clear()
        self.title_lbl.configure(text="1 / 3   Sign in")
        self._hint("Sign in and paste your API token in Chrome, "
                   "then press Auth done.").pack(fill="x", pady=(2, 7))
        row = ctk.CTkFrame(self.body, fg_color="transparent")
        row.pack(fill="x", pady=(0, 2))
        self._pill(row, "Auth done", self.do_auth_done, width=150).pack(side="left")
        self._pill(row, "Reopen", self.do_open_browser, primary=False,
                   width=100).pack(side="left", padx=10)
        self.say("waiting for you in Chrome")
        self._fit(128)

    def do_auth_done(self):
        self._clear()
        self._hint("Reading the project list...").pack(fill="x", pady=(2, 7))
        self.say("reading projects", ACCENT)
        self._fit(112)

        def job():
            bot = self.bot
            if bot is None or is_login_page(self.driver.current_url or ""):
                raise RuntimeError("Still on the login page - finish signing in.")
            snapshot_session(bot, self.cfg)
            opts, cur = list_feed_options(bot, "organization")
            if not opts:
                raise RuntimeError(
                    "No Organization dropdown found. Did the token go through?")
            return ("orgs", opts, cur)

        self._run(job)

    # ------------------------------------------------------------------ #
    # stage 3 / 4 - project and model
    # ------------------------------------------------------------------ #
    def _chooser(self, title, prompt, options, current, on_pick):
        self._clear()
        self.title_lbl.configure(text=title)
        self._hint(prompt).pack(fill="x", pady=(2, 6))
        row = ctk.CTkFrame(self.body, fg_color="transparent")
        row.pack(fill="x", pady=(0, 2))
        values = [o.get("text") or o.get("value") for o in options]
        var = ctk.StringVar(value=current if current in values else values[0])
        ctk.CTkOptionMenu(
            row, values=values, variable=var, width=360, height=29,
            corner_radius=15, fg_color=INK2, button_color=INK2,
            button_hover_color=INK3, text_color=TEXT,
            dropdown_fg_color=INK2, dropdown_text_color=TEXT,
            dropdown_hover_color=INK3,
            font=ctk.CTkFont(FONT, 12)).pack(side="left")
        self._pill(row, "Continue",
                   lambda: on_pick(next(o for o in options
                                        if (o.get("text") or o.get("value"))
                                        == var.get())),
                   width=120).pack(side="left", padx=12)
        self.say(f"{len(values)} available")
        self._fit(132)

    def stage_project(self, options, current):
        self._chooser("2 / 3   Project",
                      "Which project?",
                      options, current, self.do_pick_project)

    def do_pick_project(self, opt):
        self._clear()
        self._hint(f"Selecting {opt.get('text')} ...").pack(fill="x", pady=(2, 7))
        self.say("loading models", ACCENT)
        self._fit(112)

        def job():
            bot = self.bot
            select_feed(bot, "organization", opt["value"])
            bot.wait_idle()
            self.cfg.org = opt.get("text") or opt["value"]
            self.cfg.org_value = opt["value"]
            opts, cur = list_feed_options(bot, "schema")
            if not opts:
                raise RuntimeError("No schema dropdown found for this project.")
            return ("schemas", opts, cur)

        self._run(job)

    def stage_model(self, options, current):
        self._chooser("3 / 3   Model", "Which model?",
                      options, current, self.do_pick_model)

    def do_pick_model(self, opt):
        self._clear()
        self._hint("Loading the data...").pack(
            fill="x", pady=(4, 10))
        self.say("clicking Load Data", ACCENT)
        self._fit(112)

        def job():
            bot = self.bot
            select_feed(bot, "schema", opt["value"])
            bot.wait_idle()
            self.cfg.schema = opt.get("text") or opt["value"]
            self.cfg.schema_value = opt["value"]
            self.cfg.project = self.cfg.schema
            btn = bot.find("load_data", timeout=30, required=False)
            if btn is None:
                raise RuntimeError("Could not find the Load Data button.")
            bot.click(btn, "Load Data")
            bot.wait_idle(quiet=2.0, timeout=bot.cfg.plot_timeout)
            deadline = time.time() + 180
            while time.time() < deadline:
                if len(bot.visible_tabs()) >= 4:
                    break
                time.sleep(1)
            snapshot_session(bot, self.cfg)
            # the app fills its Date Range from the data it just loaded
            live = bot.read_date_range()
            if live:
                bot.log(f"date range on the app: {live.get('from_iso') or '?'} "
                        f"to {live.get('to_iso') or '?'}")
            else:
                bot.log("no Date Range control found after Load Data")
            return ("loaded", bot.visible_tabs(),
                    (live.get("from_iso", ""), live.get("to_iso", "")))

        self._run(job)

    # ------------------------------------------------------------------ #
    # stage 5 - the run bar
    # ------------------------------------------------------------------ #
    def stage_bar(self):
        self._clear()
        self.title_lbl.configure(text=self.cfg.schema or "Data Sanity")
        want = min(880, self._max_wl())
        if self._wl < want:
            self._apply_geometry(want, self._stage_h)
            self.update_idletasks()

        top = ctk.CTkFrame(self.body, fg_color="transparent")
        top.pack(fill="x", pady=(1, 5))

        # The boxes start EMPTY and show the app's own range as a hint. Typing
        # in them overrides the app; leaving them blank uses whatever the app
        # has for each tab and feed, which is what the report should quote.
        ctk.CTkLabel(top, text="From", text_color=MUTED,
                     font=ctk.CTkFont(FONT, 11)).pack(side="left", padx=(2, 6))
        self.e_from = self._date_box(top, 0)
        ctk.CTkLabel(top, text="to", text_color=MUTED,
                     font=ctk.CTkFont(FONT, 11)).pack(side="left", padx=8)
        self.e_to = self._date_box(top, 1)
        self.btn_dates = ctk.CTkButton(
            top, text="\u21bb", command=self.do_refresh_dates, width=30,
            height=27, corner_radius=14, fg_color=INK2, hover_color=INK3,
            text_color=TEXT, font=ctk.CTkFont(FONT, 12, "bold"))
        self.btn_dates.pack(side="left", padx=(6, 0))

        chips = ctk.CTkFrame(top, fg_color="transparent")
        chips.pack(side="left", padx=12)
        self.chips: dict[str, ctk.CTkButton] = {}
        self.chosen = {n: True for n in SECTION_ORDER}
        for name in SECTION_ORDER:
            b = ctk.CTkButton(chips, text=name, width=84, height=27,
                              corner_radius=14,
                              font=ctk.CTkFont(FONT, 10, "bold"),
                              command=lambda n=name: self._toggle_chip(n))
            b.pack(side="left", padx=3)
            self.chips[name] = b
        self._paint_chips()

        self.btn_run = self._pill(top, "Run", self.do_run, width=120)
        self.btn_run.pack(side="right", padx=(8, 2))
        self.btn_stop = self._pill(top, "Cancel", self.do_cancel,
                                   primary=False, width=100)

        self.bar = ctk.CTkProgressBar(self.body, height=4, corner_radius=2,
                                      progress_color=ACCENT, fg_color=INK2)
        if self.app_dates[0] or self.app_dates[1]:
            self.say(f"app range  {self.app_dates[0] or '?'} to "
                     f"{self.app_dates[1] or '?'}  -  leave blank to keep it")
        else:
            self.say("pick your sections and press Run")
        self._fit(140)

    def _date_box(self, parent, which: int) -> ctk.CTkEntry:
        """A date entry with a calendar button fused onto its right end."""
        box = ctk.CTkFrame(parent, fg_color=INK2, corner_radius=14, height=27)
        box.pack(side="left")
        # square entry inset past the box's rounded end, so no corner pokes out
        entry = ctk.CTkEntry(box, width=88, height=23, corner_radius=0,
                             fg_color=INK2, border_width=0, text_color=TEXT,
                             justify="center",
                             placeholder_text=self.app_dates[which] or "auto",
                             font=ctk.CTkFont(FONT, 12))
        entry.pack(side="left", padx=(10, 0), pady=2)
        ctk.CTkButton(box, text="\u25be", width=24, height=23,
                      corner_radius=12, fg_color=INK2, hover_color=INK3,
                      text_color=MUTED, font=ctk.CTkFont(FONT, 12),
                      command=lambda: self._toggle_calendar(entry, which)
                      ).pack(side="left", padx=(0, 2))
        return entry

    def _toggle_calendar(self, entry: ctk.CTkEntry, which: int) -> None:
        opened_for = self.date_pop.entry if self.date_pop else None
        self._close_calendar()
        if opened_for is not entry:
            self.date_pop = DatePopover(self, entry, self.app_dates[which],
                                        on_close=self._forget_calendar)

    def _forget_calendar(self) -> None:
        self.date_pop = None

    def _close_calendar(self) -> None:
        if self.date_pop is not None:
            self.date_pop.close()

    def _toggle_chip(self, name):
        self.chosen[name] = not self.chosen[name]
        self._paint_chips()

    def _paint_chips(self):
        for name, b in self.chips.items():
            on = self.chosen[name]
            b.configure(fg_color=ACCENT if on else INK2,
                        hover_color="#0A6FD8" if on else INK3,
                        text_color="#FFFFFF" if on else MUTED)

    def do_run(self):
        self._close_calendar()
        picked = tuple(n for n in SECTION_ORDER if self.chosen[n])
        if not picked:
            self.say("tick at least one section", WARN)
            return
        for label, val in (("From", self.e_from.get()), ("To", self.e_to.get())):
            if not val.strip():
                continue                  # blank = use the app's own range
            try:
                datetime.strptime(val.strip(), "%Y-%m-%d")
            except ValueError:
                self.say(f"{label} must be YYYY-MM-DD, or blank to use the "
                         f"app's range", BAD)
                return
        cfg = self._base_config()
        cfg.sections = picked
        cfg.date_from = self.e_from.get().strip()
        cfg.date_to = self.e_to.get().strip()
        cfg.org, cfg.schema = self.cfg.org, self.cfg.schema
        cfg.org_value = self.cfg.org_value
        cfg.schema_value = self.cfg.schema_value
        cfg.project = self.cfg.schema
        cfg.token_cache = self.cfg.token_cache
        self.bot.cfg = cfg
        self.cfg = cfg

        self.btn_run.pack_forget()
        self.btn_stop.pack(side="right", padx=(8, 2))
        self.bar.pack(fill="x", pady=(6, 2))
        self.bar.set(0)
        self._fit(152)
        self._run(lambda: run_after_load(self.driver, cfg))

    def do_refresh_dates(self):
        """Re-read the Date Range off the page, e.g. after changing feed."""
        if self.bot is None:
            self.say("open the browser first", WARN)
            return
        self.say("reading the date range off the app", ACCENT)

        def job():
            live = self.bot.read_date_range()
            return ("dates", live.get("from_iso", ""), live.get("to_iso", ""))

        self._run(job)

    def _show_app_dates(self, dfrom: str, dto: str) -> None:
        """Into the placeholders, never the values: empty still means 'auto'."""
        self.app_dates = (dfrom or "", dto or "")
        for entry, val in ((getattr(self, "e_from", None), dfrom),
                           (getattr(self, "e_to", None), dto)):
            if entry is None:
                continue
            try:
                entry.configure(placeholder_text=val or "auto")
            except Exception:
                pass
        if dfrom or dto:
            self.say(f"app range  {dfrom or '?'} to {dto or '?'}", GOOD)
        else:
            self.say("no Date Range box found on the page", WARN)

    def do_cancel(self):
        self.cancel_event.set()
        self.say("cancelling at the next checkpoint", WARN)

    # ------------------------------------------------------------------ #
    # plumbing
    # ------------------------------------------------------------------ #
    def _base_config(self) -> Config:
        return Config(
            url=DEFAULT_URL,
            date_from=self.cfg.date_from, date_to=self.cfg.date_to,
            profile_dir=DEFAULT_PROFILE,
            on_log=lambda line: self.q.put(("log", line)),
            on_progress=lambda info: self.q.put(("prog", info)),
            cancel=self.cancel_event,
        )

    def _run(self, fn):
        if self.worker and self.worker.is_alive():
            return
        self.cancel_event.clear()

        def target():
            try:
                self.q.put(("done", fn()))
            except Cancelled:
                self.q.put(("cancelled", None))
            except Exception as exc:
                self.q.put(("error", (exc, traceback.format_exc())))

        self.worker = threading.Thread(target=target, daemon=True)
        self.worker.start()

    def _drain(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self.log(payload)
                elif kind == "prog":
                    self._progress(payload)
                elif kind == "done":
                    self._done(payload)
                elif kind == "cancelled":
                    self.say("cancelled", WARN)
                    self.stage_bar()
                elif kind == "error":
                    exc, tb = payload
                    self.log(tb)
                    self.say(f"{type(exc).__name__}: {exc}", BAD)
                    if not self.log_open:
                        self._toggle_log()
                    self._recover_stage()
        except queue.Empty:
            pass
        self.after(120, self._drain)

    def _progress(self, info: dict):
        kind = info.get("kind")
        if kind == "section":
            self.say(f"section  ·  {info.get('tag')}", ACCENT)
        elif kind == "feed":
            i, n = info.get("feed_i", 0), info.get("feed_n", 1)
            self.say(f"feed {i}/{n}  ·  {info.get('tag')}", ACCENT)
            try:
                self.bar.set(i / max(1, n))
            except Exception:
                pass
        elif kind == "charts":
            self.say(f"{info.get('chart_n')} chart(s)  ·  {info.get('tag')}",
                     ACCENT)
        elif kind == "image":
            self.say(f"{info.get('tag')}  ·  chart {info.get('chart_i')}"
                     f"/{info.get('chart_n')}  ·  {info.get('images')} image(s)",
                     ACCENT)

    def _done(self, payload):
        if payload == "browser":
            self.stage_auth()
        elif isinstance(payload, tuple) and payload[:1] == ("orgs",):
            self.stage_project(payload[1], payload[2])
        elif isinstance(payload, tuple) and payload[:1] == ("schemas",):
            self.stage_model(payload[1], payload[2])
        elif isinstance(payload, tuple) and payload[:1] == ("loaded",):
            self.log(f"[ui] tabs: {payload[1]}")
            dates = payload[2] if len(payload) > 2 else ""
            if isinstance(dates, tuple):
                self.app_dates = (dates[0] or "", dates[1] or "")
            self.stage_bar()
            if self.app_dates[0] or self.app_dates[1]:
                self.say(f"loaded  -  app range {self.app_dates[0] or '?'} to "
                         f"{self.app_dates[1] or '?'}", GOOD)
            else:
                self.say("platform loaded", GOOD)
        elif isinstance(payload, tuple) and payload[:1] == ("dates",):
            self._show_app_dates(payload[1], payload[2])
        elif isinstance(payload, Path):
            self.last_output = payload
            self.stage_done(payload)
        else:
            self.say("done", GOOD)

    def stage_done(self, path: Path):
        self._clear()
        self.title_lbl.configure(text="Finished")
        self._hint(path.name, TEXT).pack(fill="x", pady=(2, 7))
        row = ctk.CTkFrame(self.body, fg_color="transparent")
        row.pack(fill="x", pady=(0, 2))
        self._pill(row, "Open report", lambda: _open(path),
                   color=GOOD).pack(side="left")
        self._pill(row, "Folder", lambda: _open(path.parent), primary=False,
                   width=100).pack(side="left", padx=10)
        self._pill(row, "New run", self.stage_bar, primary=False,
                   width=100).pack(side="left")
        self.say("report written", GOOD)
        self._fit(132)

    def _recover_stage(self):
        if self.driver is None:
            self.stage_start()
        elif self.cfg.schema:
            self.stage_bar()
        else:
            self.stage_auth()

    def _quit(self):
        if self.worker and self.worker.is_alive():
            self.cancel_event.set()
        try:
            if self.driver:
                self.driver.quit()
        except WebDriverException:
            pass
        self.destroy()


def _open(p: Path):
    try:
        if sys.platform.startswith("win"):
            os.startfile(str(p))          # noqa: S606
        elif sys.platform == "darwin":
            subprocess.run(["open", str(p)], check=False)
        else:
            subprocess.run(["xdg-open", str(p)], check=False)
    except OSError:
        pass


def main() -> int:
    Island().mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
