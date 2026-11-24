import csv
import os
import threading
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import ttk, scrolledtext, messagebox, filedialog

import requests
import tkintermapview  # pip install tkintermapview
from PIL import Image, ImageDraw, ImageTk  # installed with tkintermapview

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
# Set your key as an environment variable (recommended):
#   Windows (cmd):  setx GRAPHHOPPER_KEY "your-key-here"   (then reopen the terminal)
#   Mac/Linux:      export GRAPHHOPPER_KEY="your-key-here"
# or paste it below instead of YOUR_API_KEY_HERE.
KEY = os.environ.get("GRAPHHOPPER_KEY", "YOUR_API_KEY_HERE")

GEOCODE_URL = "https://graphhopper.com/api/1/geocode"
ROUTE_URL = "https://graphhopper.com/api/1/route"
PROFILES = ["car", "bike", "foot"]
KM_PER_MILE = 1.609344
TIMEOUT = 15  # seconds

# Map colours
COLOR_SELECTED = "#1a73e8"
COLOR_OTHER = "#9aa0a6"
COLOR_TURN_OUTER = "#d93025"
COLOR_TURN_INNER = "#ffffff"

# GraphHopper instruction "sign" -> arrow + short label
SIGN_LABELS = {
    -98: "U-turn", -8: "U-turn", -7: "Keep left", -3: "Sharp left", -2: "Left",
    -1: "Slight left", 0: "Straight", 1: "Slight right", 2: "Right",
    3: "Sharp right", 4: "Arrive", 5: "Stop", 6: "Roundabout", 7: "Keep right", 8: "U-turn",
}
SIGN_ARROWS = {
    -98: "↶", -8: "↶", -7: "↖", -3: "⬅", -2: "←", -1: "↖", 0: "↑",
    1: "↗", 2: "→", 3: "➡", 4: "⚑", 5: "●", 6: "↻", 7: "↗", 8: "↷",
}


# ---------------------------------------------------------------------------
# API helpers
# ---------------------------------------------------------------------------
def geocode(location):
    """Return (lat, lng, label) for a place name. Raises ValueError on failure."""
    r = requests.get(
        GEOCODE_URL,
        params={"q": location, "limit": 1, "key": KEY},
        timeout=TIMEOUT,
    )
    data = r.json()
    if r.status_code != 200:
        raise ValueError(f"Geocoding error ({r.status_code}): {data.get('message', 'unknown')}")
    if not data.get("hits"):
        raise ValueError(f'No results found for "{location}".')

    hit = data["hits"][0]
    parts = [hit.get("name", location), hit.get("state", ""), hit.get("country", "")]
    label = ", ".join(p for p in parts if p)
    return hit["point"]["lat"], hit["point"]["lng"], label


def fetch_routes(points, vehicle, want_alternatives):
    """Call the routing API. points is a list of (lat, lng). Returns list of paths."""
    params = [("key", KEY), ("vehicle", vehicle),
              ("points_encoded", "false")]  # plain GeoJSON coordinates, easy to draw
    for lat, lng in points:
        params.append(("point", f"{lat},{lng}"))

    if want_alternatives and len(points) == 2:
        params += [("algorithm", "alternative_route"),
                   ("alternative_route.max_paths", "3")]

    r = requests.get(ROUTE_URL, params=params, timeout=TIMEOUT)
    data = r.json()
    if r.status_code != 200:
        raise ValueError(f"Routing error ({r.status_code}): {data.get('message', 'unknown')}")
    return data["paths"]


def parse_path(path, name, vehicle, cost=None):
    """Turn an API path into drawable data: line coordinates + turn-by-turn steps."""
    coords = [(c[1], c[0]) for c in path["points"]["coordinates"]]  # GeoJSON is (lng, lat)
    steps = []
    for s in path.get("instructions", []):
        idx = min(s["interval"][0], len(coords) - 1)
        steps.append({
            "text": s["text"],
            "distance": s["distance"],
            "sign": s.get("sign", 0),
            "pos": coords[idx],  # where on the map this instruction happens
        })
    return {"name": name, "vehicle": vehicle, "coords": coords, "steps": steps,
            "distance": path["distance"], "time": path["time"], "cost": cost}


# ---------------------------------------------------------------------------
# Formatting helpers
# ---------------------------------------------------------------------------
def fmt_distance(meters, unit):
    km = meters / 1000
    return f"{km / KM_PER_MILE:.1f} mi" if unit == "miles" else f"{km:.1f} km"


def fmt_duration(ms):
    total = int(ms / 1000)
    return f"{total // 3600:02d}:{total // 60 % 60:02d}:{total % 60:02d}"


# ---------------------------------------------------------------------------
# Export helpers
# ---------------------------------------------------------------------------
def _unit_distance(meters, unit):
    km = meters / 1000
    return km / KM_PER_MILE if unit == "miles" else km


def export_csv(path, route, unit):
    """One row per instruction. utf-8-sig so Excel opens accents and symbols correctly."""
    dist_col = f"Distance ({'mi' if unit == 'miles' else 'km'})"
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(["Step", "Turn", "Instruction", dist_col, "Latitude", "Longitude"])
        for i, st in enumerate(route["steps"], start=1):
            w.writerow([i, SIGN_LABELS.get(st["sign"], ""), st["text"],
                        f"{_unit_distance(st['distance'], unit):.2f}",
                        f"{st['pos'][0]:.6f}", f"{st['pos'][1]:.6f}"])


def export_docx(path, route, routes, stops, unit):
    """Word document: trip summary, optional route overview, turn-by-turn table."""
    try:
        from docx import Document
        from docx.shared import Inches
    except ImportError:
        raise RuntimeError("Word export needs the python-docx package.\n"
                           "Install it with:  pip install python-docx")

    def add_table(doc, headers, rows, widths, header=True):
        table = doc.add_table(rows=0, cols=len(widths))
        table.style = "Table Grid"
        if header:
            cells = table.add_row().cells
            for c, h in zip(cells, headers):
                c.text = h
                for r in c.paragraphs[0].runs:
                    r.bold = True
        for row in rows:
            cells = table.add_row().cells
            for c, val in zip(cells, row):
                c.text = str(val)
        for row in table.rows:
            for c, w in zip(row.cells, widths):
                c.width = Inches(w)
        return table

    dist_label = "mi" if unit == "miles" else "km"
    doc = Document()
    doc.add_heading("Route directions", level=0)
    doc.add_paragraph(" \u2192 ".join(s[2] for s in stops))
    doc.add_paragraph(f"Generated {datetime.now():%b %d, %Y %I:%M %p}")

    doc.add_heading("Summary", level=1)
    summary = [("Route", route["name"]),
               ("Vehicle", route["vehicle"]),
               ("Distance", fmt_distance(route["distance"], unit)),
               ("Duration", fmt_duration(route["time"]))]
    if route.get("cost") is not None:
        summary.append(("Estimated fuel cost", f"\u20b1{route['cost']:,.2f}"))
    tbl = add_table(doc, None, summary, [2.0, 4.5], header=False)
    for row in tbl.rows:
        for r in row.cells[0].paragraphs[0].runs:
            r.bold = True

    if len(routes) > 1:
        doc.add_heading("All routes", level=1)
        rows = [(r["name"] + (" (shown below)" if r is route else ""), r["vehicle"],
                 fmt_distance(r["distance"], unit), fmt_duration(r["time"])) for r in routes]
        add_table(doc, ["Route", "Vehicle", "Distance", "Duration"], rows, [2.2, 1.2, 1.5, 1.6])

    doc.add_heading("Turn-by-turn directions", level=1)
    rows = [(i, SIGN_LABELS.get(st["sign"], ""), st["text"], f"{_unit_distance(st['distance'], unit):.2f} {dist_label}")
            for i, st in enumerate(route["steps"], start=1)]
    add_table(doc, ["#", "Turn", "Instruction", "Distance"], rows, [0.4, 1.1, 3.9, 1.1])
    doc.save(path)


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
class RoutePlanner(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Route Mate")
        self.geometry("1360x820")
        self.minsize(1000, 620)

        # Map state
        self.routes = []        # list of dicts from parse_path()
        self.stops = []         # list of (lat, lng, label)
        self.selected = 0       # index of the highlighted route
        self.unit = "km"
        self.highlight_marker = None

        self.turn_icon = self._make_dot(COLOR_TURN_OUTER, 14)
        self.focus_icon = self._make_dot("#fbbc04", 22)
        self._build_widgets()

    @staticmethod
    def _make_dot(color, size):
        """Small round map marker (much less bulky than the default pin)."""
        big = size * 4  # draw large, then shrink for smooth edges
        img = Image.new("RGBA", (big, big), (0, 0, 0, 0))
        d = ImageDraw.Draw(img)
        d.ellipse((0, 0, big - 1, big - 1), fill="white")
        m = big // 6
        d.ellipse((m, m, big - 1 - m, big - 1 - m), fill=color)
        return ImageTk.PhotoImage(img.resize((size, size), Image.LANCZOS))

    def _build_widgets(self):
        pad = {"padx": 8, "pady": 4}
        # Status bar first so it always stays visible at the bottom
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.status_var, anchor="w").pack(side="bottom", fill="x", padx=10, pady=(0, 6))

        # Left column = inputs + results, right = big map (drag the divider to resize)
        self.main = ttk.PanedWindow(self, orient="horizontal")
        self.main.pack(fill="both", expand=True, padx=10, pady=8)
        left = ttk.Frame(self.main, width=500)
        left.pack_propagate(False)
        self.main.add(left, weight=0)

        form = ttk.Frame(left)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)

        # Start / destination
        ttk.Label(form, text="Starting location:").grid(row=0, column=0, sticky="w", **pad)
        self.start_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.start_var).grid(row=0, column=1, columnspan=2, sticky="ew", **pad)

        ttk.Label(form, text="Destination:").grid(row=1, column=0, sticky="w", **pad)
        self.dest_var = tk.StringVar()
        ttk.Entry(form, textvariable=self.dest_var).grid(row=1, column=1, columnspan=2, sticky="ew", **pad)

        # Waypoints
        ttk.Label(form, text="Extra stops:").grid(row=2, column=0, sticky="nw", **pad)
        stop_frame = ttk.Frame(form)
        stop_frame.grid(row=2, column=1, columnspan=2, sticky="ew", **pad)
        stop_frame.columnconfigure(0, weight=1)
        self.stop_var = tk.StringVar()
        stop_entry = ttk.Entry(stop_frame, textvariable=self.stop_var)
        stop_entry.grid(row=0, column=0, sticky="ew")
        stop_entry.bind("<Return>", lambda e: self.add_stop())
        ttk.Button(stop_frame, text="Add stop", command=self.add_stop).grid(row=0, column=1, padx=(6, 0))
        ttk.Button(stop_frame, text="Remove selected", command=self.remove_stop).grid(row=0, column=2, padx=(6, 0))
        self.stop_list = tk.Listbox(stop_frame, height=2)
        self.stop_list.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(4, 0))

        # Options row
        opts = ttk.Frame(form)
        opts.grid(row=3, column=0, columnspan=3, sticky="ew", **pad)

        ttk.Label(opts, text="Vehicle:").pack(side="left")
        self.vehicle_var = tk.StringVar(value="car")
        vehicle_box = ttk.Combobox(opts, textvariable=self.vehicle_var, values=PROFILES,
                                   state="readonly", width=8)
        vehicle_box.pack(side="left", padx=(4, 16))
        vehicle_box.bind("<<ComboboxSelected>>", lambda e: self._toggle_cost_fields())

        ttk.Label(opts, text="Units:").pack(side="left")
        self.unit_var = tk.StringVar(value="km")
        ttk.Combobox(opts, textvariable=self.unit_var, values=["km", "miles"],
                     state="readonly", width=8).pack(side="left", padx=(4, 16))

        self.alt_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(form, text="Show alternative routes", variable=self.alt_var).grid(
            row=4, column=0, columnspan=3, sticky="w", **pad)

        # Cost estimate (car only)
        self.cost_frame = ttk.LabelFrame(form, text="Fuel cost estimate (car only)")
        self.cost_frame.grid(row=5, column=0, columnspan=3, sticky="ew", **pad)
        ttk.Label(self.cost_frame, text="Fuel efficiency (km/L):").pack(side="left", padx=(8, 4), pady=6)
        self.eff_var = tk.StringVar(value="12")
        self.eff_entry = ttk.Entry(self.cost_frame, textvariable=self.eff_var, width=8)
        self.eff_entry.pack(side="left")
        ttk.Label(self.cost_frame, text="Fuel price (₱/L):").pack(side="left", padx=(16, 4))
        self.price_var = tk.StringVar(value="60")
        self.price_entry = ttk.Entry(self.cost_frame, textvariable=self.price_var, width=8)
        self.price_entry.pack(side="left")

        # Buttons
        btns = ttk.Frame(form)
        btns.grid(row=6, column=0, columnspan=3, sticky="ew", **pad)
        self.go_btn = ttk.Button(btns, text="Get directions", command=self.on_submit)
        self.go_btn.pack(side="left")
        self.compare_btn = ttk.Button(btns, text="Compare modes", command=self.on_compare)
        self.compare_btn.pack(side="left", padx=6)
        ttk.Button(btns, text="Clear output", command=self.clear_output).pack(side="left", padx=6)
        self.export_btn = ttk.Button(btns, text="Export\u2026", command=self.on_export, state="disabled")
        self.export_btn.pack(side="left", padx=6)

        self.tabs = ttk.Notebook(left)
        self.tabs.pack(fill="both", expand=True, pady=(6, 0))

        self.map = tkintermapview.TkinterMapView(self.main, corner_radius=0)
        self.map.set_position(12.8797, 121.7740)  # Philippines overview until a route is loaded
        self.map.set_zoom(6)
        self.main.add(self.map, weight=1)

        # Tab 1: summary text (same as before)
        self.output = scrolledtext.ScrolledText(self.tabs, wrap="word", font=("Consolas", 10), state="disabled")
        self.tabs.add(self.output, text="Summary")

        # Tab 2: clickable turn-by-turn list
        turn_frame = ttk.Frame(self.tabs)
        self.turn_list = tk.Listbox(turn_frame, font=("Consolas", 10), activestyle="none", exportselection=False)
        scroll = ttk.Scrollbar(turn_frame, orient="vertical", command=self.turn_list.yview)
        self.turn_list.config(yscrollcommand=scroll.set)
        self.turn_list.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.turn_list.bind("<<ListboxSelect>>", self._on_turn_selected)
        self.tabs.add(turn_frame, text="Turn-by-turn")

    # --- widget helpers ---------------------------------------------------
    def _toggle_cost_fields(self):
        state = "normal" if self.vehicle_var.get() == "car" else "disabled"
        self.eff_entry.config(state=state)
        self.price_entry.config(state=state)

    def add_stop(self):
        text = self.stop_var.get().strip()
        if text:
            self.stop_list.insert("end", text)
            self.stop_var.set("")

    def remove_stop(self):
        for i in reversed(self.stop_list.curselection()):
            self.stop_list.delete(i)

    def write(self, text=""):
        self.output.config(state="normal")
        self.output.insert("end", text + "\n")
        self.output.see("end")
        self.output.config(state="disabled")

    def clear_output(self):
        self.output.config(state="normal")
        self.output.delete("1.0", "end")
        self.output.config(state="disabled")
        # Also wipe the map and turn list
        self.routes, self.stops, self.highlight_marker = [], [], None
        self.map.delete_all_path()
        self.map.delete_all_marker()
        self.turn_list.delete(0, "end")
        self.export_btn.config(state="disabled")

    # --- map drawing ------------------------------------------------------
    def show_routes(self, routes, stops, unit, fit=True):
        """Store the new result set, draw it, and zoom the map to fit it."""
        self.routes, self.stops, self.unit, self.selected = routes, stops, unit, 0
        self._redraw()
        self.export_btn.config(state="normal")
        if fit:
            self._fit_to_route(routes[self.selected]["coords"])

    def select_route(self, index):
        """Called when the user clicks a grey (alternative) route on the map."""
        if 0 <= index < len(self.routes) and index != self.selected:
            self.selected = index
            self._redraw()
            self.status_var.set(f"Selected: {self.routes[index]['name']}")

    def _fit_to_route(self, coords):
        lats = [c[0] for c in coords]
        lngs = [c[1] for c in coords]
        if len(set(lats)) == 1 and len(set(lngs)) == 1:
            self.map.set_position(lats[0], lngs[0])
            self.map.set_zoom(15)
            return
        self.map.fit_bounding_box((max(lats), min(lngs)), (min(lats), max(lngs)))

    def _redraw(self):
        self.map.delete_all_path()
        self.map.delete_all_marker()
        self.highlight_marker = None

        # Unselected routes first (so the selected one is drawn on top); click one to select it
        for i, route in enumerate(self.routes):
            if i != self.selected:
                self.map.set_path(route["coords"], color=COLOR_OTHER, width=4,
                                  command=lambda path, idx=i: self.select_route(idx))
        current = self.routes[self.selected]
        self.map.set_path(current["coords"], color=COLOR_SELECTED, width=6)

        # Start / via / end markers
        last = len(self.stops) - 1
        for k, (lat, lng, label) in enumerate(self.stops):
            title = "Start" if k == 0 else "End" if k == last else f"Stop {k}"
            self.map.set_marker(lat, lng, text=f"{title}: {label.split(',')[0]}")

        # Small red dot at every turn on the selected route; click one to jump to it in the list
        for i, step in enumerate(current["steps"]):
            if step["sign"] in (0, 4, 5):  # skip "continue straight", arrival and via points
                continue
            self.map.set_marker(step["pos"][0], step["pos"][1],
                                icon=self.turn_icon, icon_anchor="center",
                                command=self._on_turn_marker_clicked, data=i)

        # Fill the turn-by-turn list for the selected route
        self.turn_list.delete(0, "end")
        for step in current["steps"]:
            arrow = SIGN_ARROWS.get(step["sign"], "•")
            self.turn_list.insert("end", f" {arrow}  {step['text']}  ({fmt_distance(step['distance'], self.unit)})")

    def _on_turn_marker_clicked(self, marker):
        self._focus_step(marker.data, zoom=False)

    def _on_turn_selected(self, _event):
        sel = self.turn_list.curselection()
        if sel:
            self._focus_step(sel[0], zoom=True, update_list=False)

    def _focus_step(self, index, zoom=True, update_list=True):
        """Highlight one instruction on the map and in the list."""
        steps = self.routes[self.selected]["steps"] if self.routes else []
        if not (0 <= index < len(steps)):
            return
        step = steps[index]
        lat, lng = step["pos"]

        if update_list:
            self.turn_list.selection_clear(0, "end")
            self.turn_list.selection_set(index)
            self.turn_list.see(index)
            self.tabs.select(self.turn_list.master)

        if self.highlight_marker is not None:
            self.highlight_marker.delete()
        text = step["text"] if len(step["text"]) <= 45 else step["text"][:42] + "..."
        self.highlight_marker = self.map.set_marker(
            lat, lng, text=text, icon=self.focus_icon, icon_anchor="center")

        if zoom:
            self.map.set_position(lat, lng)
            self.map.set_zoom(17)
        self.status_var.set(f"{SIGN_LABELS.get(step['sign'], 'Step')}: {step['text']}")

    # --- export -----------------------------------------------------------
    def on_export(self):
        """Save the currently selected (blue) route as .docx or .csv."""
        if not self.routes:
            return
        route = self.routes[self.selected]
        default = f"directions_{route['vehicle']}_{datetime.now():%Y%m%d_%H%M}"
        path = filedialog.asksaveasfilename(
            title="Export directions", initialfile=default, defaultextension=".docx",
            filetypes=[("Word document", "*.docx"), ("CSV file (Excel)", "*.csv")])
        if not path:
            return
        ext = os.path.splitext(path)[1].lower()
        if ext not in (".docx", ".csv"):
            path += ".docx"
            ext = ".docx"
        try:
            if ext == ".csv":
                export_csv(path, route, self.unit)
            else:
                export_docx(path, route, self.routes, self.stops, self.unit)
        except PermissionError:
            messagebox.showerror("Export failed", "Can't write to that file. Is it open in Word or Excel?")
            return
        except Exception as e:
            messagebox.showerror("Export failed", str(e))
            return
        self.status_var.set(f"Saved {route['name']} to {path}")

    # --- main action ------------------------------------------------------
    def on_submit(self):
        start = self.start_var.get().strip()
        dest = self.dest_var.get().strip()
        if not start or not dest:
            messagebox.showwarning("Missing input", "Please enter both a starting location and a destination.")
            return
        if KEY == "YOUR_API_KEY_HERE":
            messagebox.showerror("API key missing", "Set your GraphHopper key in GRAPHHOPPER_KEY or in the script.")
            return

        vehicle = self.vehicle_var.get()
        cost_inputs = None
        if vehicle == "car":
            try:
                cost_inputs = (float(self.eff_var.get()), float(self.price_var.get()))
                if cost_inputs[0] <= 0 or cost_inputs[1] < 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("Invalid input", "Fuel efficiency and price must be positive numbers.")
                return

        names = [start] + list(self.stop_list.get(0, "end")) + [dest]
        job = dict(names=names, vehicle=vehicle, unit=self.unit_var.get(),
                   alternatives=self.alt_var.get(), cost=cost_inputs)

        self._set_busy(True)
        self.status_var.set("Working...")
        threading.Thread(target=self._worker, args=(job,), daemon=True).start()

    def _worker(self, job):
        """Runs in a background thread so the window doesn't freeze."""
        lines = []
        try:
            points, labels = [], []
            for name in job["names"]:
                lat, lng, label = geocode(name)
                points.append((lat, lng))
                labels.append(label)

            paths = fetch_routes(points, job["vehicle"], job["alternatives"])

            lines.append("=" * 60)
            lines.append(f"{' -> '.join(labels)}")
            lines.append(f"Vehicle: {job['vehicle']}")
            if job["alternatives"] and len(points) > 2:
                lines.append("(Alternatives only work for 2-point trips, so showing the main route.)")
            if len(paths) > 1:
                lines.append("Tip: click a grey line on the map to switch to that route.")
            lines.append("=" * 60)

            routes = []
            for i, path in enumerate(paths, start=1):
                lines.append("")
                lines.append(f"--- Route {i} of {len(paths)} ---" if len(paths) > 1 else "--- Route ---")
                lines.append(f"Distance: {fmt_distance(path['distance'], job['unit'])}")
                lines.append(f"Duration: {fmt_duration(path['time'])}")

                eta = datetime.now() + timedelta(milliseconds=path["time"])
                lines.append(f"Estimated arrival: {eta.strftime('%I:%M %p, %b %d')}")

                cost = None
                if job["cost"]:
                    km_per_l, price = job["cost"]
                    cost = (path["distance"] / 1000) / km_per_l * price
                    lines.append(f"Estimated fuel cost: ₱{cost:,.2f}")

                lines.append("")
                for step in path.get("instructions", []):
                    lines.append(f"  {step['text']} ({fmt_distance(step['distance'], job['unit'])})")

                routes.append(parse_path(path, f"Route {i}", job["vehicle"], cost))
            lines.append("")

            stops = [(p[0], p[1], lbl) for p, lbl in zip(points, labels)]
            self.after(0, self._finish, lines, "Done", (routes, stops, job["unit"]))
        except requests.exceptions.RequestException as e:
            self.after(0, self._finish, [f"Network error: {e}", ""], "Network error")
        except ValueError as e:
            self.after(0, self._finish, [f"Error: {e}", ""], "Error")
        except Exception as e:  # last-resort guard so the button always re-enables
            self.after(0, self._finish, [f"Unexpected error: {e}", ""], "Error")

    # --- route comparison -------------------------------------------------
    def _set_busy(self, busy):
        state = "disabled" if busy else "normal"
        self.go_btn.config(state=state)
        self.compare_btn.config(state=state)

    def on_compare(self):
        start = self.start_var.get().strip()
        dest = self.dest_var.get().strip()
        if not start or not dest:
            messagebox.showwarning("Missing input", "Please enter both a starting location and a destination.")
            return
        if KEY == "YOUR_API_KEY_HERE":
            messagebox.showerror("API key missing", "Set your GraphHopper key in GRAPHHOPPER_KEY or in the script.")
            return

        names = [start] + list(self.stop_list.get(0, "end")) + [dest]
        job = dict(names=names, unit=self.unit_var.get())

        self._set_busy(True)
        self.status_var.set("Comparing modes...")
        threading.Thread(target=self._compare_worker, args=(job,), daemon=True).start()

    def _compare_worker(self, job):
        """Routes the trip with every mode in PROFILES and ranks them by travel time."""
        lines = []
        try:
            points, labels = [], []
            for name in job["names"]:
                lat, lng, label = geocode(name)
                points.append((lat, lng))
                labels.append(label)

            results, skipped = [], []
            for mode in PROFILES:
                try:
                    path = fetch_routes(points, mode, False)[0]
                    results.append((mode, path["distance"], path["time"], parse_path(path, mode, mode)))
                except ValueError as e:
                    skipped.append(f"{mode}: {e}")

            results.sort(key=lambda r: r[2])  # least time first

            lines.append("=" * 60)
            lines.append(" -> ".join(labels))
            lines.append("Modes ranked from fastest to slowest")
            lines.append("Tip: click a grey line on the map to see that mode's turns.")
            lines.append("=" * 60)
            now = datetime.now()
            for rank, (mode, dist, t, _) in enumerate(results, start=1):
                eta = (now + timedelta(milliseconds=t)).strftime("%I:%M %p, %b %d")
                lines.append(f"{rank}. {mode:<6} {fmt_duration(t)}  "
                             f"{fmt_distance(dist, job['unit']):>10}  arrive {eta}")
            for msg in skipped:
                lines.append(f"Skipped {msg}")
            lines.append("")

            routes = [r[3] for r in results]
            stops = [(p[0], p[1], lbl) for p, lbl in zip(points, labels)]
            payload = (routes, stops, job["unit"]) if routes else None
            self.after(0, self._finish, lines, "Comparison done", payload)
        except requests.exceptions.RequestException as e:
            self.after(0, self._finish, [f"Network error: {e}", ""], "Network error")
        except ValueError as e:
            self.after(0, self._finish, [f"Error: {e}", ""], "Error")
        except Exception as e:  # keeps the buttons from staying disabled
            self.after(0, self._finish, [f"Unexpected error: {e}", ""], "Error")

    def _finish(self, lines, status, map_data=None):
        for line in lines:
            self.write(line)
        self.status_var.set(status)
        self._set_busy(False)
        if map_data:
            routes, stops, unit = map_data
            self.show_routes(routes, stops, unit)


if __name__ == "__main__":
    RoutePlanner().mainloop()
