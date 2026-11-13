# DevNet-Projects: GraphHopper Route Planner

A desktop route planner built with Python and Tkinter. It started as the Lab 4.9.2
directions app and was extended by Team Trifecta (Nicdao, Pelias, Sergio) for
Project Activity 3. It uses the GraphHopper Geocoding and Routing APIs.

## What it does

You type a starting point and a destination (and optional stops), pick a vehicle,
and the app shows the distance, travel time, estimated arrival time and
turn-by-turn directions.

### Features

| Feature | How it works |
|---|---|
| Extra stops | Add stopovers with the **Add stop** button. Every place becomes one more point in the routing request. |
| Fuel cost | For cars only. Enter km/L and fuel price (defaults 12 km/L and 60 per liter). Cost = distance in km / km per liter x price. It is an estimate. |
| ETA | Estimated arrival = current time + the trip duration returned by the API. |
| Route optimization | The **Compare modes** button routes the same trip with every mode in `PROFILES` (car, bike, foot) and ranks them from fastest to slowest, with distance and arrival time. |
| Simple UI | A Tkinter window with fields, buttons, a scrollable output box and a status bar. Network calls run in background threads so the window does not freeze. |
| Also included | km or miles toggle, car/bike/foot dropdown, and up to 3 alternative routes (see limitations). |

## Requirements

- Python 3 (Tkinter comes with the standard Windows and macOS installers)
- The `requests` package: `pip install requests`
- A free GraphHopper API key: https://www.graphhopper.com/

## Setup and run

1. Get an API key from GraphHopper.
2. Save it in an environment variable. Do not paste it into the code or commit it.
   - Windows (cmd): `setx GRAPHHOPPER_KEY "your-key-here"`, then open a new terminal
   - Mac/Linux: `export GRAPHHOPPER_KEY="your-key-here"`
3. Run the app: `python route_planner.py`

If the key is missing, the app shows an "API key missing" message instead of calling the API.

## How to use

1. Enter a starting location and a destination.
2. (Optional) Type a stop and press **Add stop** or Enter. Use **Remove selected** to delete one.
3. Pick a vehicle and units. For a car, check the fuel efficiency and price.
4. Press **Get directions** for the route, or **Compare modes** to rank car, bike and foot.

## Code map (`route_planner.py`)

| Section | What is in it |
|---|---|
| Config | API URLs, `KEY`, `PROFILES`, unit constant, request timeout |
| API helpers | `geocode()` turns a place name into latitude/longitude. `fetch_routes()` sends all points to the Routing API and returns the paths. |
| Formatting helpers | `fmt_distance()` and `fmt_duration()` |
| GUI: `RoutePlanner` | `_build_widgets()` lays out the window. `on_submit()` / `_worker()` handle **Get directions**. `on_compare()` / `_compare_worker()` handle **Compare modes**. `_finish()` prints results on the main thread. |

How a request flows: button press -> validate input -> background thread ->
`geocode()` for each place -> `fetch_routes()` -> format lines -> `self.after(0, ...)`
hands the text back to the main thread.

## How to continue the project

- **Add a vehicle:** add its name to `PROFILES`. The dropdown and Compare modes pick it up.
  If your GraphHopper plan does not support the vehicle, Compare modes skips it and says why.
- **Add an option:** create the widget in `_build_widgets()`, read it in `on_submit()`,
  put it in the `job` dictionary, and use it in `_worker()`.
- **Keep the window responsive:** never call the API from a button handler directly. Use a thread
  like `_worker()` and update the screen only with `self.after(0, ...)`.

## Known limitations

- Alternative routes only work for trips with exactly two points (no extra stops). With stops, the app shows the main route and says so.
- Fuel cost is an estimate and only applies to cars. The currency symbol is fixed to the peso sign.
- The free GraphHopper plan has request limits and may not support every vehicle type.
- Each place name is geocoded to its top search result, so ambiguous names may pick the wrong place.

## Backlog (future enhancements)

- A map view of the route
- Save or export directions to a file
- A more visually appealing UI
- Automated tests

## Team

Team Trifecta: Andrea Pelias (Scrum Leader and Recorder), Kyleen Nicdao (Dev/Tech Lead), Juliana Sergio (Dev/QA Lead).
