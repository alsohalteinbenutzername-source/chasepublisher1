from __future__ import annotations

import argparse
from datetime import datetime, timezone
import json
import logging
import math
import os
import time
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, urlsplit
from urllib.request import Request, urlopen
from zoneinfo import ZoneInfo


LOG = logging.getLogger("stormchase-daily-outlook-publisher")
BERLIN = ZoneInfo("Europe/Berlin")
OUTLOOK_EXTENT = (4.0, 16.0, 47.0, 56.0)
GRID_SPACING = 0.4
EXPECTED_GRID_POINTS = 744
RETRY_ATTEMPTS = 4


class PublisherError(RuntimeError):
    pass


def _configuration() -> tuple[str, str]:
    url = os.environ.get("SUPABASE_URL", "").strip().rstrip("/")
    key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY", "").strip()
    parsed = urlsplit(url)
    if parsed.scheme != "https" or not parsed.netloc or parsed.path not in {"", "/"}:
        raise PublisherError("SUPABASE_URL must be an HTTPS Supabase project URL.")
    if parsed.username or parsed.password or parsed.query or parsed.fragment:
        raise PublisherError("SUPABASE_URL must not contain credentials, paths, or a query.")
    if not key:
        raise PublisherError("SUPABASE_SERVICE_ROLE_KEY is not configured.")
    return url, key


def _request(
    base_url: str,
    service_key: str,
    method: str,
    path: str,
    *,
    query: dict[str, str] | None = None,
    payload: Any = None,
) -> Any:
    url = f"{base_url}{path}"
    if query:
        url = f"{url}?{urlencode(query)}"
    headers = {
        "apikey": service_key,
        "Authorization": f"Bearer {service_key}",
        "Accept": "application/json",
    }
    body = None
    if payload is not None:
        body = json.dumps(payload, separators=(",", ":"), allow_nan=False).encode("utf-8")
        headers["Content-Type"] = "application/json"
    if method == "POST":
        headers["Prefer"] = "resolution=merge-duplicates,return=minimal"

    for attempt in range(RETRY_ATTEMPTS):
        try:
            request = Request(url, data=body, headers=headers, method=method)
            with urlopen(request, timeout=60) as response:
                contents = response.read()
                content_type = response.headers.get("Content-Type", "")
            if not contents:
                return None
            if "application/json" not in content_type.lower():
                raise PublisherError(f"Supabase returned non-JSON data for {method} {path}.")
            try:
                return json.loads(contents)
            except json.JSONDecodeError as error:
                raise PublisherError("Supabase returned malformed JSON.") from error
        except HTTPError as error:
            detail = error.read().decode("utf-8", errors="replace")
            retryable = error.code in {408, 425, 429} or error.code >= 500
            if not retryable or attempt + 1 >= RETRY_ATTEMPTS:
                raise PublisherError(
                    f"Supabase {method} {path} returned HTTP {error.code}: {detail[:500]}"
                ) from error
            retry_after = error.headers.get("Retry-After")
            try:
                delay = min(max(float(retry_after), 1.0), 60.0) if retry_after else 2 ** (attempt + 1)
            except ValueError:
                delay = 2 ** (attempt + 1)
            time.sleep(delay)
        except (URLError, TimeoutError, OSError) as error:
            if attempt + 1 >= RETRY_ATTEMPTS:
                raise PublisherError(f"Could not reach Supabase: {error}") from error
            time.sleep(2 ** (attempt + 1))
    raise PublisherError(f"Supabase request failed: {method} {path}.")


def _axis_points(start: float, end: float) -> list[float]:
    values = [start + index * GRID_SPACING for index in range(int((end - start) / GRID_SPACING) + 1)]
    if values[-1] < end:
        values.append(end)
    return values


def _grid_points() -> list[tuple[float, float]]:
    west, east, south, north = OUTLOOK_EXTENT
    points = [
        (latitude, longitude)
        for latitude in _axis_points(south, north)
        for longitude in _axis_points(west, east)
    ]
    if len(points) != EXPECTED_GRID_POINTS:
        raise PublisherError(f"Expected {EXPECTED_GRID_POINTS} outlook points, got {len(points)}.")
    return points


def _validate_grid(grid: Any) -> None:
    if not isinstance(grid, list) or len(grid) != EXPECTED_GRID_POINTS:
        raise PublisherError(f"Outlook payload must contain exactly {EXPECTED_GRID_POINTS} points.")
    for index, point in enumerate(grid):
        if not isinstance(point, dict) or not isinstance(point.get("icon_d2"), dict):
            raise PublisherError(f"Outlook point {index} is missing ICON-D2 hourly data.")
        hourly = point["icon_d2"]
        times = hourly.get("time")
        if not isinstance(times, list) or not times:
            raise PublisherError(f"Outlook point {index} has no hourly timestamps.")
        if len(times) != len(set(times)):
            raise PublisherError(f"Outlook point {index} has duplicate timestamps.")
        if not isinstance(hourly.get("cin"), list) or len(hourly["cin"]) != len(times):
            raise PublisherError(f"Outlook point {index} has invalid CIN values.")
        fields = (
            "cape", "convective_inhibition", "temperature_850hPa", "temperature_700hPa",
            "temperature_500hPa", "temperature_2m", "dew_point_2m", "relative_humidity_2m",
            "precipitation", "snowfall", "wind_speed_10m", "wind_direction_10m",
            "wind_speed_925hPa", "wind_direction_925hPa", "wind_speed_850hPa",
            "wind_direction_850hPa", "wind_speed_700hPa", "wind_direction_700hPa",
            "wind_speed_500hPa", "wind_direction_500hPa",
        )
        for field in fields:
            values = hourly.get(field)
            if not isinstance(values, list) or len(values) != len(times):
                raise PublisherError(f"Outlook point {index} has invalid {field} values.")
            for value in values:
                if value is not None and (
                    isinstance(value, bool)
                    or not isinstance(value, (int, float))
                    or not math.isfinite(value)
                ):
                    raise PublisherError(f"Outlook point {index} has non-finite {field} values.")


def _fetch_daily_grid() -> list[dict[str, Any]]:
    # Reuse the app's validated Open-Meteo request, pacing, retry, and CIN logic.
    from weather_api import get_germany_grid_weather

    grid = get_germany_grid_weather(_grid_points())
    _validate_grid(grid)
    return grid


def publish_daily_outlook(*, force: bool = False) -> None:
    base_url, service_key = _configuration()
    forecast_date = datetime.now(BERLIN).date()
    date_text = forecast_date.isoformat()
    rows = _request(
        base_url,
        service_key,
        "GET",
        "/rest/v1/daily_outlook_forecasts",
        query={
            "select": "forecast_date,grid_payload",
            "forecast_date": f"eq.{date_text}",
            "limit": "1",
        },
    )
    if not isinstance(rows, list) or not all(isinstance(row, dict) for row in rows):
        raise PublisherError("Supabase returned an invalid daily-outlook response.")
    if rows and rows[0].get("forecast_date") == date_text and not force:
        try:
            _validate_grid(rows[0].get("grid_payload"))
        except PublisherError:
            LOG.warning("Today's stored outlook is incomplete; refetching it.")
        else:
            LOG.info("Today's outlook is already stored; skipping the API fetch.")
            return

    LOG.info("Fetching the 744-point ICON-D2 outlook for %s.", date_text)
    grid = _fetch_daily_grid()
    issued_at = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
    _request(
        base_url,
        service_key,
        "POST",
        "/rest/v1/daily_outlook_forecasts",
        query={"on_conflict": "forecast_date"},
        payload=[{
            "forecast_date": date_text,
            "issued_at": issued_at,
            "grid_payload": grid,
        }],
    )
    LOG.info("Stored today's ICON-D2 outlook (%d points), issued at %s.", len(grid), issued_at)


def main() -> None:
    parser = argparse.ArgumentParser(description="Publish StormChase's daily ICON-D2 outlook.")
    parser.add_argument(
        "--force",
        action="store_true",
        help="refetch and replace today's snapshot (intended for manual runs)",
    )
    arguments = parser.parse_args()
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )
    if os.environ.get("GITHUB_EVENT_NAME") == "schedule":
        local_now = datetime.now(BERLIN)
        if local_now.hour != 0:
            LOG.info(
                "Skipping UTC schedule at %s; it is not midnight in Europe/Berlin.",
                local_now.isoformat(),
            )
            return
    publish_daily_outlook(force=arguments.force)


if __name__ == "__main__":
    main()
