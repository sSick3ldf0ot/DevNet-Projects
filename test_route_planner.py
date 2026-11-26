
import pytest
from unittest.mock import Mock
import route_planner



# A. DISTANCE AND TIME FORMATTING (4 TESTS)


def test_format_distance_kilometers():
    assert route_planner.fmt_distance(5000, "km") == "5.0 km"


def test_format_distance_miles():
    assert route_planner.fmt_distance(1609.344, "miles") == "1.0 mi"


def test_format_duration():
    assert route_planner.fmt_duration(3661000) == "01:01:01"


def test_format_zero_duration():
   result = route_planner.fmt_duration(0)
   assert result == "00:00:00"



# B. GEOCODING (3 TESTS)


def test_geocode_success(monkeypatch):
    fake_response = Mock()
    fake_response.status_code = 200
    fake_response.json.return_value = {
        "hits": [{
            "name": "Manila",
            "state": "Metro Manila",
            "country": "Philippines",
            "point": {
                "lat": 14.5995,
                "lng": 120.9842
            }
        }]
    }

    monkeypatch.setattr(
        route_planner.requests,
        "get",
        lambda *args, **kwargs: fake_response
    )

    result = route_planner.geocode("Manila")

    assert result == (
        14.5995,
        120.9842,
        "Manila, Metro Manila, Philippines"
    )


def test_geocode_no_results(monkeypatch):
    fake_response = Mock()
    fake_response.status_code = 200
    fake_response.json.return_value = {"hits": []}

    monkeypatch.setattr(
        route_planner.requests,
        "get",
        lambda *args, **kwargs: fake_response
    )

    with pytest.raises(ValueError, match="No results found"):
        route_planner.geocode("Unknown Location")


def test_geocode_api_error(monkeypatch):
    fake_response = Mock()
    fake_response.status_code = 400
    fake_response.json.return_value = {
        "message": "Invalid API request"
    }

    monkeypatch.setattr(
        route_planner.requests,
        "get",
        lambda *args, **kwargs: fake_response
    )

    with pytest.raises(ValueError, match="Geocoding error"):
        route_planner.geocode("Manila")



# C. ROUTING (3 TESTS)

def test_routing_intermediate_stops(monkeypatch):
    fake_response = Mock()
    fake_response.status_code = 200
    fake_response.json.return_value = {
        "paths": [{"distance": 1000, "time": 5000}]
    }

    captured = {}

    def fake_get(url, params, timeout):
        captured["params"] = params
        return fake_response

    monkeypatch.setattr(
        route_planner.requests,
        "get",
        fake_get
    )

    points = [
        (14.5995, 120.9842),
        (14.6760, 121.0437),
        (14.5547, 121.0244)
    ]

    route_planner.fetch_routes(points, "car", False)

    submitted_points = [
        value for key, value in captured["params"]
        if key == "point"
    ]

    assert submitted_points == [
        "14.5995,120.9842",
        "14.676,121.0437",
        "14.5547,121.0244"
    ]


def test_routing_alternatives_only_for_two_points(monkeypatch):
    fake_response = Mock()
    fake_response.status_code = 200
    fake_response.json.return_value = {
        "paths": [{"distance": 1000, "time": 5000}]
    }

    captured_params = []

    def fake_get(url, params, timeout):
        captured_params.append(params)
        return fake_response

    monkeypatch.setattr(
        route_planner.requests,
        "get",
        fake_get
    )

    two_points = [
        (14.5995, 120.9842),
        (14.5547, 121.0244)
    ]

    three_points = [
        (14.5995, 120.9842),
        (14.6760, 121.0437),
        (14.5547, 121.0244)
    ]

    route_planner.fetch_routes(two_points, "car", True)
    route_planner.fetch_routes(three_points, "car", True)

    first_request = captured_params[0]
    second_request = captured_params[1]

    assert ("algorithm", "alternative_route") in first_request
    assert ("algorithm", "alternative_route") not in second_request


def test_routing_api_error(monkeypatch):
    fake_response = Mock()
    fake_response.status_code = 400
    fake_response.json.return_value = {
        "message": "Invalid routing request"
    }

    monkeypatch.setattr(
        route_planner.requests,
        "get",
        lambda *args, **kwargs: fake_response
    )

    points = [
        (14.5995, 120.9842),
        (14.5547, 121.0244)
    ]

    with pytest.raises(ValueError, match="Routing error"):
        route_planner.fetch_routes(points, "car", False)


def test_export_csv(tmp_path):
    route = {
        "steps": [{
            "sign": 2,
            "text": "Turn right",
            "distance": 1500,
            "pos": (14.5995, 120.9842)
        }]
    }

    file_path = tmp_path / "directions.csv"

    route_planner.export_csv(
        str(file_path),
        route,
        "km"
    )

    assert file_path.exists()

    content = file_path.read_text(encoding="utf-8-sig")

    assert "Turn right" in content
    assert "1.50" in content
    assert "Latitude" in content
    assert "Longitude" in content
    assert "14.5995" in content
    assert "120.9842" in content