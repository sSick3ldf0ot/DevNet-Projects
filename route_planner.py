import os
import threading
import tkinter as tk
from datetime import datetime, timedelta
from tkinter import ttk, scrolledtext, messagebox

import requests

# ---------------------------------------------------------------------------
# Config
# ---------------------------------------------------------------------------
# Set your key as an environment variable (recommended):
#   Windows (cmd):  setx GRAPHHOPPER_KEY "your-key-here"
#   Mac/Linux:      export GRAPHHOPPER_KEY="your-key-here"
# or paste it below instead of YOUR_API_KEY_HERE.
KEY = "09e599e9-232b-4660-8b6e-ac91c1233e03"

GEOCODE_URL = "https://graphhopper.com/api/1/geocode"
ROUTE_URL = "https://graphhopper.com/api/1/route"
PROFILES = ["car", "bike", "foot"]
KM_PER_MILE = 1.609344
TIMEOUT = 15  # seconds


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
    params = [("key", KEY), ("vehicle", vehicle)]
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
# GUI
# ---------------------------------------------------------------------------
class RoutePlanner(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("GraphHopper Route Planner")
        self.geometry("760x780")
        self.minsize(640, 600)
        self._build_widgets()

    def _build_widgets(self):
        pad = {"padx": 8, "pady": 4}
        form = ttk.Frame(self)
        form.pack(fill="x", padx=10, pady=8)
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
        self.stop_list = tk.Listbox(stop_frame, height=4)
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
        ttk.Checkbutton(opts, text="Show alternative routes", variable=self.alt_var).pack(side="left")

        # Cost estimate (car only)
        self.cost_frame = ttk.LabelFrame(form, text="Fuel cost estimate (car only)")
        self.cost_frame.grid(row=4, column=0, columnspan=3, sticky="ew", **pad)
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
        btns.grid(row=5, column=0, columnspan=3, sticky="ew", **pad)
        self.go_btn = ttk.Button(btns, text="Get directions", command=self.on_submit)
        self.go_btn.pack(side="left")
        ttk.Button(btns, text="Clear output", command=self.clear_output).pack(side="left", padx=6)

        # Output
        self.output = scrolledtext.ScrolledText(self, wrap="word", font=("Consolas", 10), state="disabled")
        self.output.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(self, textvariable=self.status_var, anchor="w").pack(fill="x", padx=10, pady=(0, 6))

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

        self.go_btn.config(state="disabled")
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
            lines.append("=" * 60)

            for i, path in enumerate(paths, start=1):
                lines.append("")
                lines.append(f"--- Route {i} of {len(paths)} ---" if len(paths) > 1 else "--- Route ---")
                lines.append(f"Distance: {fmt_distance(path['distance'], job['unit'])}")
                lines.append(f"Duration: {fmt_duration(path['time'])}")

                eta = datetime.now() + timedelta(milliseconds=path["time"])
                lines.append(f"Estimated arrival: {eta.strftime('%I:%M %p, %b %d')}")

                if job["cost"]:
                    km_per_l, price = job["cost"]
                    cost = (path["distance"] / 1000) / km_per_l * price
                    lines.append(f"Estimated fuel cost: ₱{cost:,.2f}")

                lines.append("")
                for step in path.get("instructions", []):
                    lines.append(f"  {step['text']} ({fmt_distance(step['distance'], job['unit'])})")
            lines.append("")
            self.after(0, self._finish, lines, "Done")
        except requests.exceptions.RequestException as e:
            self.after(0, self._finish, [f"Network error: {e}", ""], "Network error")
        except ValueError as e:
            self.after(0, self._finish, [f"Error: {e}", ""], "Error")
        except Exception as e:  # last-resort guard so the button always re-enables
            self.after(0, self._finish, [f"Unexpected error: {e}", ""], "Error")

    def _finish(self, lines, status):
        for line in lines:
            self.write(line)
        self.status_var.set(status)
        self.go_btn.config(state="normal")


if __name__ == "__main__":
    RoutePlanner().mainloop()
