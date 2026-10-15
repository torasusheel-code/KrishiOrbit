from flask import (
    Flask, render_template, request, redirect,
    url_for, session, flash, jsonify, send_file, Response
)
import sqlite3
import os
from datetime import date
import requests
from dotenv import load_dotenv
import time
import json
import subprocess
from functools import wraps
from werkzeug.utils import secure_filename
import uuid
import io
import csv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from collections import defaultdict
load_dotenv(
    dotenv_path=os.path.join(os.path.dirname(__file__), ".env"),
    override=True
)

app = Flask(__name__)

DATA_GOV_API_KEY = os.getenv("DATA_GOV_API_KEY")

CROP_PRODUCTION_API_KEY = os.getenv("CROP_PRODUCTION_API_KEY")

# ---------------- ADMIN CONFIGURATION ----------------

ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD")

app.secret_key = os.environ.get("SECRET_KEY")

# ---------------- DATABASE CONFIGURATION ----------------

DATABASE = os.path.join(app.root_path, "KrishiOrbit.db")

# ---------------- FARMER IMAGE UPLOAD CONFIGURATION ----------------

UPLOAD_FOLDER = os.path.join(
    app.root_path, "static", "uploads"
)

os.makedirs(UPLOAD_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {"jpg", "jpeg", "png", "webp"}

MAX_IMAGE_SIZE = 5 * 1024 * 1024

app.config["MAX_CONTENT_LENGTH"] = 6 * 1024 * 1024

# ---------------- WEATHER CONFIGURATION ----------------

WEATHER_CACHE = {}
WEATHER_CACHE_SECONDS = 1800

# ---------------- CROP PRODUCTION API ----------------

CROP_PRODUCTION_RESOURCE_ID = (
    "35be999b-0208-4354-b557-f6ca9a5355de"
)

CROP_PRODUCTION_API_URL = (
    "https://api.data.gov.in/resource/"
    + CROP_PRODUCTION_RESOURCE_ID
)

CROP_PRODUCTION_API_KEY = os.getenv(
    "CROP_PRODUCTION_API_KEY"
)
# ---------------- EXTENDED CROP INFORMATION ----------------

CROP_EXTRA_INFO = {
    "wheat": {
        "soil": "Well-drained loamy soil is generally suitable.",
        "sowing_time": "Rabi season; timing depends on the region.",
        "irrigation": "Irrigation needs depend on soil, rainfall and crop stage.",
        "care": "Monitor weeds, pests and diseases regularly."
    },
    "rice": {
        "soil": "Soil suitability depends on the rice variety and local conditions.",
        "sowing_time": "Usually grown during Kharif in many Indian regions.",
        "irrigation": "Water management depends on the cultivation method.",
        "care": "Monitor weeds, pests and diseases regularly."
    },
    "maize": {
        "soil": "Well-drained soil is generally preferred.",
        "sowing_time": "Often grown during Kharif; regional seasons vary.",
        "irrigation": "Water requirements depend on weather and crop stage.",
        "care": "Monitor weeds, pests and diseases regularly."
    }
}

# -------- CROP PRODUCTION EXPORT HELPER --------

def get_crop_production_export_records(filters):
    """
    Fetch matching crop production records for CSV/Excel export.
    Uses the same government API as the existing dashboard.
    """

    if not CROP_PRODUCTION_API_KEY:
        raise ValueError("Crop production API key is not configured.")

    records = []
    offset = 0
    limit = 100
    max_pages = 100

    for _ in range(max_pages):
        params = {
            "api-key": CROP_PRODUCTION_API_KEY,
            "format": "json",
            "limit": limit,
            "offset": offset
        }

        api_filters = {
            "filters[state_name]": filters.get("state", "").strip(),
            "filters[district_name]": filters.get("district", "").strip(),
            "filters[crop]": filters.get("crop", "").strip(),
            "filters[season]": filters.get("season", "").strip(),
            "filters[crop_year]": filters.get("year", "").strip()
        }

        params.update({
            key: value
            for key, value in api_filters.items()
            if value
        })

        data = fetch_crop_production_data(params)

        if not isinstance(data, dict):
            raise ValueError("Unexpected government API response.")

        if data.get("error") or data.get("status") == "error":
            raise ValueError("Government API returned an error.")

        batch = data.get("records", [])

        if not isinstance(batch, list):
            raise ValueError("Invalid records received from API.")

        if not batch:
            break

        records.extend(
            record for record in batch
            if isinstance(record, dict)
        )

        offset += len(batch)

        total = int(data.get("total", 0))

        if offset >= total or len(batch) < limit:
            break

    return records
# -------- CROP PRODUCTION CHARTS API --------

@app.route("/api/crop-production/charts")
def crop_production_charts_api():

    filters = {
        "state": request.args.get("state", "").strip(),
        "district": request.args.get("district", "").strip(),
        "crop": request.args.get("crop", "").strip(),
        "season": request.args.get("season", "").strip(),
        "year": request.args.get("year", "").strip()
    }

    if filters["year"] and not filters["year"].isdigit():
        return jsonify({
            "success": False,
            "message": "Year must be a valid number."
        }), 400

    try:
        records = get_crop_production_export_records(filters)

        crop_totals = defaultdict(float)
        state_totals = defaultdict(float)
        season_totals = defaultdict(float)
        yearly_totals = defaultdict(float)

        total_area = 0.0

        for item in records:
            crop_name = str(item.get("crop") or "Unknown")
            state_name = str(item.get("state_name") or "Unknown")
            season_name = str(item.get("season") or "Unknown")
            year_value = item.get("crop_year")

            production_value = item.get("production_")
            area_value = item.get("area_")

            # Calculate total cultivated area
            try:
                area = float(area_value)

                if area >= 0:
                    total_area += area

            except (ValueError, TypeError):
                pass

            # Aggregate production
            try:
                production = float(production_value)

                if production < 0:
                    continue

            except (ValueError, TypeError):
                continue

            crop_totals[crop_name] += production
            state_totals[state_name] += production
            season_totals[season_name] += production

            if year_value not in (None, ""):
                try:
                    year_label = str(int(float(year_value)))
                    yearly_totals[year_label] += production
                except (ValueError, TypeError):
                    pass

        # Sort chart values by production
        def sorted_items(values):
            return [
                {
                    "label": label,
                    "production": round(value, 3)
                }
                for label, value in sorted(
                    values.items(),
                    key=lambda pair: pair[1],
                    reverse=True
                )
            ]

        return jsonify({
            "success": True,
            "source": "data.gov.in",
            "record_count": len(records),
            "total_area": round(total_area, 3),

            "charts": {
                "crop_wise": sorted_items(crop_totals),
                "state_wise": sorted_items(state_totals),
                "season_wise": sorted_items(season_totals),

                "year_wise": [
                    {
                        "label": year,
                        "production": round(
                            yearly_totals[year], 3
                        )
                    }
                    for year in sorted(
                        yearly_totals,
                        key=int
                    )
                ]
            }
        })

    except (requests.RequestException, subprocess.TimeoutExpired) as e:
        app.logger.exception(
            "Crop production charts API request failed."
        )

        return jsonify({
            "success": False,
            "message": "Unable to fetch data for charts.",
            "error_type": type(e).__name__,
            "error_details": str(e)[:500]
        }), 502

    except (ValueError, TypeError) as e:
        app.logger.exception(
            "Crop production charts API failed."
        )

        return jsonify({
            "success": False,
            "message": "Unable to process chart data.",
            "error_type": type(e).__name__,
            "error_details": str(e)[:500]
        }), 502

    except Exception:
        app.logger.exception(
            "Unexpected crop production charts error."
        )

        return jsonify({
            "success": False,
            "message": "An unexpected error occurred while loading charts."
        }), 500

# -------- CROP PRODUCTION MAP API --------

@app.route("/api/crop-production/map")
def crop_production_map_api():
    filters = {
        "state": "",
        "district": "",
        "crop": request.args.get("crop", "").strip(),
        "season": request.args.get("season", "").strip(),
        "year": request.args.get("year", "").strip()
    }

    if filters["year"] and not filters["year"].isdigit():
        return jsonify({
            "success": False,
            "message": "Invalid year."
        }), 400

    try:
        records = get_crop_production_export_records(filters)

        state_totals = defaultdict(float)

        for item in records:
            state_name = str(
                item.get("state_name") or ""
            ).strip()

            production_value = item.get("production_")

            if not state_name:
                continue

            try:
                production = float(production_value)

                if production >= 0:
                    state_totals[state_name] += production

            except (ValueError, TypeError):
                continue

        return jsonify({
            "success": True,
            "source": "data.gov.in",
            "state_production": [
                {
                    "state": state,
                    "production": round(total, 3)
                }
                for state, total in sorted(
                    state_totals.items(),
                    key=lambda item: item[1],
                    reverse=True
                )
            ]
        })

    except (requests.RequestException, subprocess.TimeoutExpired,
            ValueError, TypeError) as e:
        app.logger.exception(
            "Crop production map API failed."
        )

        return jsonify({
            "success": False,
            "message": "Unable to load state-wise production.",
            "error_type": type(e).__name__
        }), 502

# ---------------- CROP PRODUCTION API HELPER ----------------

def fetch_crop_production_data(params):
    """
    Fetch crop production data from data.gov.in.

    Windows:
        Uses curl.exe.

    Linux / Render:
        Uses requests.
    """

    if os.name == "nt":

        # Build curl command
        command = [
            "curl.exe",
            "-sS",
            "--fail",
            "--retry", "2",
            "--retry-delay", "1",
            "--connect-timeout", "10",
            "--max-time", "30",
            "-G",
            CROP_PRODUCTION_API_URL,
        ]

        # Add API parameters
        for key, value in params.items():
            command.extend([
                "--data-urlencode",
                f"{key}={value}",
            ])

        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                timeout=35,
                check=False,
            )

        except subprocess.TimeoutExpired as exc:
            app.logger.error(
                "Crop production API curl timed out."
            )
            raise requests.Timeout(
                "Crop production API request timed out."
            ) from exc

        # Check curl errors
        if result.returncode != 0:
            app.logger.error(
                "Crop API curl failed. Return code: %s",
                result.returncode,
            )

            app.logger.error(
                "Crop API curl stderr: %s",
                result.stderr[:1000],
            )

            app.logger.error(
                "Crop API curl stdout: %s",
                result.stdout[:500],
            )

            raise requests.RequestException(
                f"Crop production curl failed with code "
                f"{result.returncode}: {result.stderr[:300]}"
            )

        # Parse JSON response
        try:
            return json.loads(result.stdout)

        except json.JSONDecodeError as exc:
            app.logger.error(
                "Crop production API returned invalid JSON."
            )
            raise ValueError(
                "Invalid JSON from crop production API."
            ) from exc

    # ---------------- LINUX / RENDER ----------------

    response = requests.get(
        CROP_PRODUCTION_API_URL,
        params=params,
        timeout=(10, 20),
    )

    response.raise_for_status()

    try:
        return response.json()

    except ValueError as exc:
        app.logger.error(
            "Crop production API returned invalid JSON."
        )
        raise ValueError(
            "Invalid JSON from crop production API."
        ) from exc
@app.route("/api/crop-production/connection-test")
def crop_production_connection_test():
    try:
        response = requests.get(
            CROP_PRODUCTION_API_URL,
            params={
                "api-key": CROP_PRODUCTION_API_KEY,
                "format": "json",
                "limit": 1
            },
            timeout=(10, 15)
        )

        return jsonify({
            "success": response.ok,
            "http_status": response.status_code,
            "message": (
                "Government API responded."
                if response.ok
                else "Government API returned an HTTP error."
            )
        }), 200

    except requests.RequestException as exc:
        app.logger.exception(
            "Government API connection test failed."
        )

        return jsonify({
            "success": False,
            "error_type": type(exc).__name__,
            "error_details": str(exc)[:500],
            "message": "Could not connect to Government API."
        }), 502
# -------- CROP PRODUCTION CSV DOWNLOAD --------

@app.route("/api/crop-production/download/csv")
def download_crop_production_csv():

    filters = {
        "state": request.args.get("state", ""),
        "district": request.args.get("district", ""),
        "crop": request.args.get("crop", ""),
        "season": request.args.get("season", ""),
        "year": request.args.get("year", "")
    }

    if filters["year"] and not filters["year"].isdigit():
        return jsonify({
            "success": False,
            "message": "Year must be a valid number."
        }), 400

    try:
        records = get_crop_production_export_records(filters)

        output = io.StringIO(newline="")

        writer = csv.writer(output)

        writer.writerow([
            "Crop Year",
            "State",
            "District",
            "Crop",
            "Season",
            "Area",
            "Production"
        ])

        for item in records:
            writer.writerow([
                item.get("crop_year", ""),
                item.get("state_name", ""),
                item.get("district_name", ""),
                item.get("crop", ""),
                item.get("season", ""),
                item.get("area_", ""),
                item.get("production_", "")
            ])

        csv_content = "\ufeff" + output.getvalue()

        return Response(
            csv_content,
            mimetype="text/csv; charset=utf-8",
            headers={
                "Content-Disposition":
                    "attachment; filename=crop_production.csv"
            }
        )

    except (requests.RequestException, subprocess.TimeoutExpired):
        app.logger.exception("CSV export API request failed.")

        return jsonify({
            "success": False,
            "message": "Unable to fetch data for CSV export."
        }), 502

    except (ValueError, TypeError):
        app.logger.exception("CSV export failed.")

        return jsonify({
            "success": False,
            "message": "Unable to prepare CSV file."
        }), 502

# -------- CROP PRODUCTION EXCEL DOWNLOAD --------

@app.route("/api/crop-production/download/excel")
def download_crop_production_excel():

    filters = {
        "state": request.args.get("state", ""),
        "district": request.args.get("district", ""),
        "crop": request.args.get("crop", ""),
        "season": request.args.get("season", ""),
        "year": request.args.get("year", "")
    }

    if filters["year"] and not filters["year"].isdigit():
        return jsonify({
            "success": False,
            "message": "Year must be a valid number."
        }), 400

    try:
        records = get_crop_production_export_records(filters)

        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Crop Production"

        headers = [
            "Crop Year",
            "State",
            "District",
            "Crop",
            "Season",
            "Area",
            "Production"
        ]

        sheet.append(headers)

        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(
                fill_type="solid",
                fgColor="176B3A"
            )
            cell.alignment = Alignment(horizontal="center")

        for item in records:
            sheet.append([
                item.get("crop_year", ""),
                item.get("state_name", ""),
                item.get("district_name", ""),
                item.get("crop", ""),
                item.get("season", ""),
                item.get("area_", ""),
                item.get("production_", "")
            ])

        # Set readable column widths
        widths = {
            "A": 14,
            "B": 25,
            "C": 25,
            "D": 28,
            "E": 18,
            "F": 16,
            "G": 18
        }

        for column, width in widths.items():
            sheet.column_dimensions[column].width = width

        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions

        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="crop_production.xlsx",
            mimetype=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            )
        )

    except (requests.RequestException, subprocess.TimeoutExpired):
        app.logger.exception("Excel export API request failed.")

        return jsonify({
            "success": False,
            "message": "Unable to fetch data for Excel export."
        }), 502

    except (ValueError, TypeError):
        app.logger.exception("Excel export failed.")

        return jsonify({
            "success": False,
            "message": "Unable to prepare Excel file."
        }), 502
# ---------------- CROP PRODUCTION DATA API ----------------

@app.route("/api/crop-production")
def crop_production_api():

    app.logger.info(
        "Crop production API key configured: %s",
        bool(CROP_PRODUCTION_API_KEY)
    )

    # Check API key
    if not CROP_PRODUCTION_API_KEY:
        return {
            "success": False,
            "message": "Crop production API key is not configured."
        }, 503

    # Read filters
    state = request.args.get("state", "").strip()
    district = request.args.get("district", "").strip()
    crop = request.args.get("crop", "").strip()
    season = request.args.get("season", "").strip()
    year = request.args.get("year", "").strip()

    # Validate limit and offset
    try:
        limit = int(request.args.get("limit", 10))
        offset = int(request.args.get("offset", 0))

    except ValueError:
        return {
            "success": False,
            "message": "Limit and offset must be valid numbers."
        }, 400

    if not 1 <= limit <= 100:
        return {
            "success": False,
            "message": "Limit must be between 1 and 100."
        }, 400

    if offset < 0:
        return {
            "success": False,
            "message": "Offset cannot be negative."
        }, 400

    if year and not year.isdigit():
        return {
            "success": False,
            "message": "Year must be a valid number."
        }, 400

    # Build API parameters
    params = {
        "api-key": CROP_PRODUCTION_API_KEY,
        "format": "json",
        "limit": limit,
        "offset": offset,
        "sort[crop_year]": "desc"
    }

    filters = {
        "filters[state_name]": state,
        "filters[district_name]": district,
        "filters[crop]": crop,
        "filters[season]": season,
        "filters[crop_year]": year
    }

    params.update({
        key: value
        for key, value in filters.items()
        if value
    })
    app.logger.info(
    "Crop filters received: state=%r, district=%r, crop=%r, season=%r, year=%r",
    state, district, crop, season, year
)

    # Fetch data and return response
    try:
        data = fetch_crop_production_data(params)

        if not isinstance(data, dict):
            return {
                "success": False,
                "message": "Unexpected response from crop production API."
            }, 502

        if data.get("error") or data.get("status") == "error":
            app.logger.error(
                "Government crop API returned an error."
            )

            return {
                "success": False,
                "message": "Government API returned an error."
            }, 502

        records = data.get("records")

        if not isinstance(records, list):
            return {
                "success": False,
                "message": "Crop production records are unavailable."
            }, 502

        return {
            "success": True,
            "source": "data.gov.in",
            "total": int(data.get("total", 0)),
            "count": len(records),
            "limit": limit,
            "offset": offset,
            "filters": {
                "state": state,
                "district": district,
                "crop": crop,
                "season": season,
                "year": year
            },
            "records": records
        }

    except subprocess.TimeoutExpired:
        app.logger.warning(
            "Crop production API command timed out."
        )

        return {
            "success": False,
            "message": "Crop production API timed out. Please try again."
        }, 504

    except requests.Timeout:
        app.logger.warning(
            "Crop production API request timed out."
        )

        return {
            "success": False,
            "message": "Crop production API timed out. Please try again."
        }, 504

    except requests.RequestException:
        app.logger.exception(
            "Crop production API request failed."
        )

        return {
            "success": False,
            "message": "Unable to fetch crop production data right now."
        }, 502

    except (ValueError, TypeError):
        app.logger.exception(
            "Invalid crop production API response."
        )

        return {
            "success": False,
            "message": "Crop production API returned invalid data."
        }, 502
# ---------------- CROP PRODUCTION DISTRICT OPTIONS API ----------------

@app.route("/api/crop-production/districts")
def crop_production_districts_api():

    state = request.args.get("state", "").strip()

    if not state:
        return {
            "success": False,
            "message": "Please select a state first.",
            "districts": []
        }, 400

    if not CROP_PRODUCTION_API_KEY:
        return {
            "success": False,
            "message": "Crop production API key is not configured.",
            "districts": []
        }, 503

    districts = set()
    page_size = 100
    max_pages = 20

    try:
        for page in range(max_pages):

            params = {
                "api-key": CROP_PRODUCTION_API_KEY,
                "format": "json",
                "limit": page_size,
                "offset": page * page_size,
                "filters[state_name]": state
            }

            app.logger.info(
                "Fetching districts: state=%s, page=%s, offset=%s",
                state,
                page + 1,
                page * page_size
            )

            data = fetch_crop_production_data(params)

            if not isinstance(data, dict):
                return {
                    "success": False,
                    "message": "Unexpected API response.",
                    "districts": []
                }, 502

            if data.get("error") or data.get("status") == "error":
                app.logger.error(
                    "Government API returned an error during district lookup."
                )

                return {
                    "success": False,
                    "message": "Government API returned an error.",
                    "districts": []
                }, 502

            records = data.get("records", [])

            if not isinstance(records, list):
                return {
                    "success": False,
                    "message": "District records are unavailable.",
                    "districts": []
                }, 502

            for record in records:
                if not isinstance(record, dict):
                    continue

                name = record.get("district_name")

                if name and str(name).strip():
                    districts.add(str(name).strip())

            if len(records) < page_size:
                break

        return {
            "success": True,
            "state": state,
            "districts": sorted(districts),
            "count": len(districts)
        }

    except subprocess.TimeoutExpired:
        app.logger.exception("District lookup timed out.")

        return {
            "success": False,
            "message": "District lookup timed out. Please try again.",
            "districts": []
        }, 504

    except requests.Timeout:
        app.logger.exception("District lookup timed out.")

        return {
            "success": False,
            "message": "District lookup timed out. Please try again.",
            "districts": []
        }, 504

    except requests.RequestException as e:
        app.logger.exception(
            "District lookup failed: %s", str(e)
        )

        return {
            "success": False,
            "message": f"District API request failed: {str(e)}",
            "districts": []
        }, 502

    except (ValueError, TypeError):
        app.logger.exception(
            "Invalid district API response."
        )

        return {
            "success": False,
            "message": "District API returned invalid data.",
            "districts": []
        }, 502

 # -------- CROP PRODUCTION AVAILABLE YEARS API --------

@app.route("/api/crop-production/years")
def crop_production_years_api():

    state = request.args.get("state", "").strip()
    district = request.args.get("district", "").strip()
    crop = request.args.get("crop", "").strip()
    season = request.args.get("season", "").strip()

    # Cache verified from previous API tests
    state_years = {
        "punjab": list(range(2014, 1996, -1)),
        "odisha": list(range(2015, 1996, -1))
    }

    if state.lower() in state_years:
        years = state_years[state.lower()]

        return jsonify({
            "success": True,
            "years": years,
            "count": len(years),
            "filters": {
                "state": state,
                "district": district,
                "crop": crop,
                "season": season
            }
        })

    # Fetch actual available years for other states
    try:
        years_found = set()
        offset = 0
        page_size = 100
        max_pages = 10

        for _ in range(max_pages):

            params = {
                "api-key": DATA_GOV_API_KEY,
                "format": "json",
                "limit": page_size,
                "offset": offset
            }

            if state:
                params["filters[state_name]"] = state

            if district:
                params["filters[district_name]"] = district

            if crop:
                params["filters[crop]"] = crop

            if season:
                params["filters[season]"] = season

            result = fetch_crop_production_data(params)

            records = result.get("records", [])

            if not records:
                break

            for record in records:
                year = record.get("crop_year")

                if year not in (None, ""):
                    try:
                        years_found.add(int(float(year)))
                    except (ValueError, TypeError):
                        pass

            if len(records) < page_size:
                break

            offset += page_size

        years = sorted(years_found, reverse=True)

        if not years:
            return jsonify({
                "success": False,
                "message": "No available years found for these filters.",
                "years": []
            }), 404

        return jsonify({
            "success": True,
            "years": years,
            "count": len(years),
            "filters": {
                "state": state,
                "district": district,
                "crop": crop,
                "season": season
            }
        })

    except (requests.RequestException, subprocess.TimeoutExpired) as exc:
        app.logger.warning(
            "Could not fetch available crop years: %s", exc
        )

        return jsonify({
            "success": False,
            "message": "Government API is temporarily unavailable.",
            "years": []
        }), 502

    except (ValueError, TypeError) as exc:
        app.logger.exception("Error processing crop years.")

        return jsonify({
            "success": False,
            "message": "Could not process available years.",
            "years": []
        }), 502
    
    # ---------------- CROP PRODUCTION DASHBOARD ----------------

@app.route("/crop-production")
def crop_production_page():
    return render_template("crop_production.html")
# ---------------- CROP INFORMATION API ----------------

@app.route("/api/crop-info/<crop_name>")
def crop_info_api(crop_name):
    conn = get_db()

    try:
        crop = conn.execute(
            """
            SELECT id, name, season, duration, description
            FROM crops
            WHERE LOWER(name) = LOWER(?)
            """,
            (crop_name.strip(),)
        ).fetchone()

        if crop is None:
            return {
                "success": False,
                "message": "Crop not found."
            }, 404

        crop_details = {
            "id": crop["id"],
            "name": crop["name"],
            "season": crop["season"],
            "duration": crop["duration"],
            "description": crop["description"]
        }

        extra_info = CROP_EXTRA_INFO.get(
            crop["name"].strip().lower(),
            {}
        )

        crop_details.update(extra_info)

        return {
            "success": True,
            "source": "SQLite",
            "crop": crop_details
        }

    except sqlite3.Error:
        app.logger.exception("Crop information database error")

        return {
            "success": False,
            "message": "Unable to fetch crop information."
        }, 500

    finally:
        conn.close()


# ---------------- DATABASE ----------------

def get_db():
    print("GET_DB DATABASE PATH:", os.path.abspath(DATABASE))

    conn = sqlite3.connect(DATABASE)
    conn.row_factory = sqlite3.Row
    return conn
# ---------------- CROP SOWING CALENDAR ----------------
def get_sowing_recommendation(crop, state, district):

    conn = get_db()

    try:
        row = conn.execute("""
            SELECT
                sowing_start,
                sowing_end,
                source,
                recommended_period
            FROM sowing_calendar
            WHERE LOWER(TRIM(crop_name)) = LOWER(TRIM(?))
              AND LOWER(TRIM(state)) = LOWER(TRIM(?))
              AND (
                  district IS NULL
                  OR TRIM(district) = ''
                  OR LOWER(TRIM(district)) = LOWER(TRIM(?))
              )
              AND verified = 1
            ORDER BY
                CASE
                    WHEN LOWER(TRIM(district)) = LOWER(TRIM(?))
                    THEN 0
                    ELSE 1
                END
            LIMIT 1
        """, (crop, state, district, district)).fetchone()

    finally:
        conn.close()

    # No verified record found
    if row is None:
        return {
            "status": "unavailable",
            "recommended_period": "Information unavailable",
            "message": (
                f"Verified sowing dates for {crop} in "
                f"{district}, {state} are not available yet. "
                "Please consult your local agriculture department."
            ),
            "source": None,
            "verified": False
        }

    # Validate dates
    if not row["sowing_start"] or not row["sowing_end"]:
        return {
            "status": "unavailable",
            "recommended_period": "Information unavailable",
            "message": (
                "Verified sowing dates are not available yet. "
                "Please consult your local agriculture department."
            ),
            "source": row["source"],
            "verified": False
        }

    # Parse dates
    try:
        today = date.today()
        start = date.fromisoformat(row["sowing_start"])
        end = date.fromisoformat(row["sowing_end"])

    except (ValueError, TypeError):
        return {
            "status": "unavailable",
            "recommended_period": "Information unavailable",
            "message": "The saved sowing dates are invalid.",
            "source": row["source"],
            "verified": False
        }

    # Check date range
    if start > end:
        return {
            "status": "unavailable",
            "recommended_period": "Information unavailable",
            "message": "The sowing calendar dates need correction.",
            "source": row["source"],
            "verified": False
        }

    # Determine sowing status
    if start <= today <= end:
        status = "period_active"
        message = (
            "Today falls within the recorded sowing period. "
            "Confirm current local conditions before sowing."
        )

    elif today < start:
        status = "period_not_started"
        message = "The recorded sowing period has not started yet."

    else:
        status = "period_ended"
        message = "The recorded sowing period has ended."

    period = (row["recommended_period"] or "").strip()

    return {
        "status": status,
        "recommended_period": (
            period if period
            else f"{start:%d %b %Y} – {end:%d %b %Y}"
        ),
        "message": message,
        "source": row["source"],
        "verified": True
    }

# ---------------- DATABASE INITIALIZATION ----------------

def init_db():
    conn = sqlite3.connect(DATABASE)
    print("INIT_DB DATABASE PATH:", os.path.abspath(DATABASE))
    cursor = conn.cursor()

    try:
        # 1. Create crops table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS crops (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                season TEXT NOT NULL,
                duration TEXT NOT NULL,
                description TEXT
            )
        """)

        # 2. Create sowing calendar table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS sowing_calendar (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                crop_name TEXT NOT NULL,
                state TEXT NOT NULL,
                district TEXT,
                variety TEXT,
                sowing_start TEXT,
                sowing_end TEXT,
                source TEXT,
                verified INTEGER DEFAULT 0,
                recommended_period TEXT
            )
        """)

        # 3. Create crop diseases table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS crop_diseases (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                crop_name TEXT NOT NULL,
                disease_name TEXT NOT NULL,
                symptoms TEXT NOT NULL,
                prevention TEXT NOT NULL,
                verified INTEGER DEFAULT 0
            )
        """)

        # 4. Create farmer requests table
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS farmer_requests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                farmer_name TEXT,
                mobile TEXT,
                state TEXT NOT NULL,
                district TEXT NOT NULL,
                crop_name TEXT NOT NULL,
                problem_type TEXT NOT NULL,
                description TEXT NOT NULL,
                image_filename TEXT,
                status TEXT NOT NULL DEFAULT 'Pending',
                admin_reply TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)

        # 5. Add source column to crop_diseases if missing
        disease_columns = {
            row[1]
            for row in cursor.execute(
                "PRAGMA table_info(crop_diseases)"
            ).fetchall()
        }

        if "source" not in disease_columns:
            cursor.execute("""
                ALTER TABLE crop_diseases
                ADD COLUMN source TEXT
            """)

        # 6. Add recommended_period column if missing
        sowing_columns = {
            row[1]
            for row in cursor.execute(
                "PRAGMA table_info(sowing_calendar)"
            ).fetchall()
        }

        if "recommended_period" not in sowing_columns:
            cursor.execute("""
                ALTER TABLE sowing_calendar
                ADD COLUMN recommended_period TEXT
            """)

        # 7. Insert Rajouri Maize record if missing
        cursor.execute("""
            SELECT COUNT(*)
            FROM sowing_calendar
            WHERE LOWER(crop_name) = 'maize'
              AND LOWER(state) = 'jammu'
              AND LOWER(district) = 'rajouri'
        """)

        if cursor.fetchone()[0] == 0:
            cursor.execute("""
                INSERT INTO sowing_calendar (
                    crop_name,
                    state,
                    district,
                    variety,
                    sowing_start,
                    sowing_end,
                    source,
                    verified,
                    recommended_period
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                "Maize",
                "Jammu",
                "Rajouri",
                None,
                "2026-04-01",
                "2026-05-31",
                "Jammu Agriculture Department - District-wise Sowing Schedule",
                0,
                "Not verified"
            ))

        # 8. Crop master data
        crops = [
            (
                "Wheat",
                "Rabi",
                "110-150 days",
                "Wheat is an important cereal crop used to make flour."
            ),
            (
                "Rice",
                "Kharif",
                "90-150 days",
                "Rice is a major food crop that requires sufficient water for growth."
            ),
            (
                "Maize",
                "Kharif",
                "80-120 days",
                "Maize is used for both human food and animal feed."
            ),
            (
                "Mustard",
                "Rabi",
                "100-150 days",
                "Mustard is an important oilseed crop grown for its seeds and oil."
            ),
            (
                "Potato",
                "Season varies by region",
                "70-120 days",
                "Potatoes grow best in suitable soil with good drainage."
            ),
            (
                "Tomato",
                "Season varies by region",
                "90-150 days",
                "Tomato is a vegetable crop used in many dishes."
            ),
            (
                "Onion",
                "Rabi / Kharif",
                "100-150 days",
                "Onion is an important vegetable used in everyday cooking."
            ),
            (
                "Gram",
                "Rabi",
                "90-120 days",
                "Gram is a pulse crop and an important source of protein."
            ),
            (
                "Peas",
                "Rabi",
                "60-90 days",
                "Peas are a cool-season vegetable and pulse crop."
            ),
            (
                "Soybean",
                "Kharif",
                "90-120 days",
                "Soybean is an oilseed and protein-rich legume crop."
            ),
            (
                "Bajra",
                "Kharif",
                "75-100 days",
                "Bajra is a drought-tolerant millet crop that grows well in warm conditions."
            )
        ]

        # 9. Insert only crops that do not already exist
        for crop in crops:
            cursor.execute("""
                SELECT COUNT(*)
                FROM crops
                WHERE LOWER(TRIM(name)) = LOWER(TRIM(?))
            """, (crop[0],))

            if cursor.fetchone()[0] == 0:
                cursor.execute("""
                    INSERT INTO crops (
                        name,
                        season,
                        duration,
                        description
                    )
                    VALUES (?, ?, ?, ?)
                """, crop)

        # 10. Commit all database changes
        conn.commit()

        print("Database initialized successfully.")
        print("Crop data checked successfully.")

    except Exception as e:
        conn.rollback()
        print("Database initialization error:", e)
        raise

    finally:
        conn.close()


# ---------------- HOME PAGE ----------------

@app.route("/")
def home():
    return render_template("index.html")


# ---------------- CROP GUIDE ----------------

@app.route("/crops")
def crops_page():
    conn = get_db()

    crops = conn.execute("""
        SELECT * FROM crops ORDER BY name
    """).fetchall()

    conn.close()

    return render_template(
        "crops.html",
        crops=crops
    )


# ================= CROP DISEASE ADVISORY =================

@app.route("/disease-advisory", methods=["GET", "POST"])
def disease_advisory():

    results = []
    selected_crop = ""
    selected_symptom = ""

    image_result = None
    image_error = None

    symptom_keywords = {
        "spots": [
            "spot", "spots", "blight", "blast",
            "pustule", "pustules", "rust"
        ],
        "yellowing": [
            "yellow", "yellowing", "chlorotic"
        ],
        "wilting": [
            "wilt", "wilting"
        ],
        "drying": [
            "dry", "drying", "dried"
        ]
    }

    if request.method == "POST":

        selected_crop = request.form.get(
            "crop", ""
        ).strip()

        selected_symptom = request.form.get(
            "symptom", ""
        ).strip().lower()

        if selected_crop and selected_symptom:

            conn = get_db()

            try:
                rows = conn.execute("""
                    SELECT disease_name, symptoms,
                           prevention, source, verified
                    FROM crop_diseases
                    WHERE LOWER(crop_name) = LOWER(?)
                """, (selected_crop,)).fetchall()

                keywords = symptom_keywords.get(
                    selected_symptom,
                    [selected_symptom]
                )

                for row in rows:
                    item = dict(row)

                    description = (
                        (item["symptoms"] or "") + " " +
                        (item["disease_name"] or "")
                    ).lower()

                    if any(
                        keyword in description
                        for keyword in keywords
                    ):
                        results.append(item)

            finally:
                conn.close()

    return render_template(
        "disease_advisory.html",
        results=results,
        selected_crop=selected_crop,
        selected_symptom=selected_symptom,
        image_result=image_result,
        image_error=image_error
    )
# ---------------- WEATHER ----------------

from datetime import datetime
import time

WEATHER_CACHE = {}
WEATHER_CACHE_SECONDS = 1800  # 30 minutes
@app.route("/weather", methods=["GET", "POST"])
def weather():

    city = request.values.get("city", "Shimla").strip()
    if not city:
        city = "Shimla"

    weather_data = None
    forecast = []
    hourly_forecast = []
    farming_advice = []
    rain_alert = None
    irrigation_advice = None
    error = None

    try:
        # ----------------------------------
        # STEP 1: Find City Coordinates
        # ----------------------------------
        geo_url = "https://geocoding-api.open-meteo.com/v1/search"

        geo_response = requests.get(
            geo_url,
            params={
                "name": city,
                "count": 1,
                "language": "en",
                "format": "json"
            },
            timeout=15
        )

        geo_response.raise_for_status()
        geo_data = geo_response.json()

        locations = geo_data.get("results", [])

        if not locations:
            raise ValueError(
                f"'{city}' was not found. Please enter a valid city name."
            )

        location = locations[0]

        latitude = location["latitude"]
        longitude = location["longitude"]

        city = location["name"]
        country = location.get("country", "")

        # ----------------------------------
               
        # STEP 2: Get Weather Data with Cache
        # ----------------------------------

        cache_key = f"{latitude},{longitude}"
        cached_weather = WEATHER_CACHE.get(cache_key)

        if (
            cached_weather
            and time.time() - cached_weather["saved_at"]
            < WEATHER_CACHE_SECONDS
        ):
            data = cached_weather["data"]

        else:
            weather_url = "https://api.open-meteo.com/v1/forecast"

            weather_response = requests.get(
                weather_url,
                params={
                    "latitude": latitude,
                    "longitude": longitude,
                    "current": [
                        "temperature_2m",
                        "relative_humidity_2m",
                        "apparent_temperature",
                        "precipitation",
                        "wind_speed_10m",
                        "weather_code"
                    ],
                    "hourly": [
                        "temperature_2m",
                        "precipitation_probability",
                        "precipitation",
                        "relative_humidity_2m",
                        "wind_speed_10m",
                        "weather_code"
                    ],
                    "daily": [
                        "temperature_2m_max",
                        "temperature_2m_min",
                        "precipitation_probability_max",
                        "precipitation_sum",
                        "weather_code",
                        "sunrise",
                        "sunset"
                    ],
                    "forecast_days": 7,
                    "timezone": "auto"
                },
                timeout=20
            )

            weather_response.raise_for_status()
            data = weather_response.json()

            WEATHER_CACHE[cache_key] = {
                "data": data,
                "saved_at": time.time()
            }
        # ----------------------------------
        # STEP 3: Current Weather
        # ----------------------------------
        current = data.get("current", {})

        temperature = current.get("temperature_2m")
        humidity = current.get("relative_humidity_2m")
        feels_like = current.get("apparent_temperature")
        precipitation = current.get("precipitation")
        wind = current.get("wind_speed_10m")
        current_code = current.get("weather_code")

        weather_data = {
            "city": city,
            "country": country,
            "temperature": temperature,
            "humidity": humidity,
            "feels_like": feels_like,
            "precipitation": precipitation,
            "wind": wind,
            "weather_code": current_code
        }

        # ----------------------------------
        # STEP 4: Hourly Forecast
        # ----------------------------------
        hourly = data.get("hourly", {})

        hourly_times = hourly.get("time", [])
        hourly_temps = hourly.get("temperature_2m", [])
        hourly_rain_chance = hourly.get(
            "precipitation_probability", []
        )
        hourly_rainfall = hourly.get("precipitation", [])
        hourly_humidity = hourly.get("relative_humidity_2m", [])
        hourly_wind = hourly.get("wind_speed_10m", [])
        hourly_codes = hourly.get("weather_code", [])

        # Current local time from the API
        current_time = current.get("time", "")

        # Find the current hour in the hourly forecast
        start_index = 0

        for i, time_value in enumerate(hourly_times):
            if time_value >= current_time:
                start_index = i
                break

        for i in range(start_index, min(start_index + 12,
                                         len(hourly_times))):

            time_value = hourly_times[i]

            parsed_time = datetime.strptime(
                time_value, "%Y-%m-%dT%H:%M"
            )

            hourly_forecast.append({
                "time": parsed_time.strftime("%I %p"),
                "temperature": (
                    hourly_temps[i]
                    if i < len(hourly_temps) else None
                ),
                "rain_chance": (
                    hourly_rain_chance[i]
                    if i < len(hourly_rain_chance) else None
                ),
                "rainfall": (
                    hourly_rainfall[i]
                    if i < len(hourly_rainfall) else None
                ),
                "humidity": (
                    hourly_humidity[i]
                    if i < len(hourly_humidity) else None
                ),
                "wind": (
                    hourly_wind[i]
                    if i < len(hourly_wind) else None
                ),
                "weather_code": (
                    hourly_codes[i]
                    if i < len(hourly_codes) else None
                )
            })

        # ----------------------------------
        # STEP 5: Seven-Day Forecast
        # ----------------------------------
        daily = data.get("daily", {})

        dates = daily.get("time", [])
        max_temps = daily.get("temperature_2m_max", [])
        min_temps = daily.get("temperature_2m_min", [])
        rain_chances = daily.get(
            "precipitation_probability_max", []
        )
        rainfall_values = daily.get("precipitation_sum", [])
        weather_codes = daily.get("weather_code", [])
        sunrises = daily.get("sunrise", [])
        sunsets = daily.get("sunset", [])

        for i, forecast_date in enumerate(dates):

            parsed_date = datetime.strptime(
                forecast_date, "%Y-%m-%d"
            )

            forecast.append({
                "date": forecast_date,
                "day_name": parsed_date.strftime("%a"),
                "formatted_date": parsed_date.strftime("%d %b"),

                "max_temp": (
                    max_temps[i]
                    if i < len(max_temps) else None
                ),

                "min_temp": (
                    min_temps[i]
                    if i < len(min_temps) else None
                ),

                "rain_chance": (
                    rain_chances[i]
                    if i < len(rain_chances) else None
                ),

                "rainfall": (
                    rainfall_values[i]
                    if i < len(rainfall_values) else None
                ),

                "weather_code": (
                    weather_codes[i]
                    if i < len(weather_codes) else None
                ),

                "sunrise": (
                    sunrises[i][-5:]
                    if i < len(sunrises) else "--"
                ),

                "sunset": (
                    sunsets[i][-5:]
                    if i < len(sunsets) else "--"
                )
            })

        # ----------------------------------
        # STEP 6: Rain Alert
        # ----------------------------------
        upcoming_rain = [
            item["rain_chance"]
            for item in hourly_forecast
            if item["rain_chance"] is not None
        ]

        max_rain_chance = (
            max(upcoming_rain) if upcoming_rain else None
        )

        if max_rain_chance is not None:

            if max_rain_chance >= 70:
                rain_alert = {
                    "level": "high",
                    "title": "High Rain Possibility",
                    "message": (
    f"There is a maximum {max_rain_chance}% chance of rain "
    "in the next 12 hours. Check the local forecast and "
    "soil moisture before irrigating."
)
                }

            elif max_rain_chance >= 40:
                rain_alert = {
                    "level": "medium",
                    "title": "Possible Rain",
                   "message": (
    f"There is a maximum {max_rain_chance}% chance of rain "
    "in the next 12 hours. Check crop requirements and "
    "soil moisture before deciding to irrigate."
)
                }

            else:
                rain_alert = {
                    "level": "low",
                    "title": "Low Rain Possibility",
                   "message": (
    f"There is a maximum {max_rain_chance}% chance of rain "
    "in the next 12 hours. Check crop requirements and "
    "soil moisture before deciding to irrigate."
)
                }

        # ----------------------------------
        # STEP 7: Smart Irrigation Advice
        # ----------------------------------
        next_rainfall = [
            item["rainfall"]
            for item in hourly_forecast
            if item["rainfall"] is not None
        ]

        expected_rainfall = sum(next_rainfall)

        if max_rain_chance is not None and max_rain_chance >= 70:

            irrigation_advice = {
                "title": "Rain-Aware Irrigation",
               # High chance of rain
            
               "message": (
    "Rain is likely. Check soil moisture, crop water "
    "requirements, and actual rainfall before irrigating."
               ),

                "icon": "🌧️"
            }

        elif expected_rainfall >= 5:

            irrigation_advice = {
                "title": "Check Soil Moisture",
                "message": (
    f"Approximately {expected_rainfall:.1f} mm of rainfall "
    "is forecast in the next 12 hours. Check soil moisture "
    "before irrigating."
                ),
                "icon": "💧"
            }

        elif humidity is not None and humidity < 35:

            irrigation_advice = {
                "title": "Check Water Requirement",
               "message": (
    "Humidity is low. Consider crop requirements, "
    "temperature, and soil moisture before deciding "
    "whether irrigation is needed."

                ),
                "icon": "🌱"
            }

        else:

            irrigation_advice = {
                "title": "Irrigation Check",
                "message": (
    "Check soil moisture, crop growth stage, and local "
    "weather conditions before deciding to irrigate."

                ),
                "icon": "🌿"
            }

                # ----------------------------------
        # STEP 8: Farming Advice
        # ----------------------------------
        if temperature is not None and temperature >= 35:
            farming_advice.append(
                "🌡️ Temperatures are high. Monitor crops for heat "
                "stress and check soil moisture."
            )

        if temperature is not None and temperature <= 10:
            farming_advice.append(
                "❄️ Temperatures are low. Follow local agricultural "
                "guidance to protect sensitive crops from cold."
            )

        if humidity is not None and humidity >= 80:
            farming_advice.append(
                "🌿 Humidity is high. Monitor crops for signs of "
                "disease and maintain proper field ventilation."
            )

        if max_rain_chance is not None and max_rain_chance >= 70:
            farming_advice.append(
                "🌧️ Rain is likely. Check the local forecast when "
                "planning harvesting, fertilizer application, "
                "or spraying."
            )

        if not farming_advice:
            farming_advice.append(
                "🌱 Monitor your crops regularly and irrigate "
                "according to soil moisture and crop requirements."
            )

    except ValueError as e:
        error = str(e)

    except requests.RequestException:
        app.logger.exception("Weather API request failed")
        error = (
            "Unable to retrieve weather data. Please check "
            "your internet connection and try again."
        )

    except (KeyError, TypeError, IndexError):
        app.logger.exception("Unable to process weather data")
        error = (
            "Unable to process weather data. Please try again later."
        )

    return render_template(
        "weather.html",
        weather=weather_data,
        forecast=forecast,
        hourly_forecast=hourly_forecast,
        city=city,
        error=error,
        farming_advice=farming_advice,
        rain_alert=rain_alert,
        irrigation_advice=irrigation_advice
    )
        
        
# ---------------- SMART SOWING ADVISOR API ----------------


@app.route("/api/sowing-advice", methods=["POST"])
def sowing_advice():

    payload = request.get_json(silent=True) or {}

    crop = str(payload.get("crop", "")).strip().lower()
    state = str(payload.get("state", "")).strip().lower()
    district = str(payload.get("district", "")).strip()

    allowed_states = {
        "jammu": "Jammu",
        "punjab": "Punjab",
        "himachal": "Himachal Pradesh",
        "himachal pradesh": "Himachal Pradesh"
    }

    # STEP 1: VALIDATE INPUT
    if not crop or not district or state not in allowed_states:
        return {
            "success": False,
            "message": "Please select a supported state, district and crop."
        }, 400

    state_name = allowed_states[state]
    location_query = f"{district}, {state_name}, India"
        # Known districts in Jammu region
    known_districts = {
        "jammu", "samba", "kathua", "udhampur",
        "reasi", "rajouri", "poonch", "doda",
        "kishtwar", "ramban"
    }

        # Known district coordinates for Jammu region
    district_coordinates = {
        "samba": {
            "latitude": 32.56,
            "longitude": 75.12
        }
    }

    if state_name == "Jammu" and district.strip().lower() not in known_districts:
        return {
            "success": False,
            "message": "Please enter a valid Jammu region district."
        }, 400
    
    try:

        # STEP 2: FIND LOCATION
        geo_response = requests.get(
            "https://geocoding-api.open-meteo.com/v1/search",
            params={
                "name": district,
                "count": 10,
                "language": "en",
                "format": "json",
                "countryCode": "IN"
            },
            timeout=10
        )

        geo_response.raise_for_status()

        results = geo_response.json().get("results", [])

        print("District searched:", district)
        print("Geocoding results:", results)

               # STEP 3: VERIFY DISTRICT AND STATE
        district_lower = district.strip().lower()
        location = None

        for item in results:

            place_name = item.get("name", "").strip().lower()
            admin1 = item.get("admin1", "").strip().lower()
            admin2 = item.get("admin2", "").strip().lower()
            country = item.get("country_code", "").strip().upper()

            # Check district name
            district_matches = (
                place_name == district_lower
                or admin2 == district_lower
                or district_lower in place_name
            )

            if not district_matches:
                continue

            # Confirm India
            if country and country != "IN":
                continue

            # Verify state/region
            if state_name == "Jammu":
                state_matches = (
                    "jammu and kashmir" in admin1
                    or "jammu" in admin1
                    or "jammu" in admin2
                )

            elif state_name == "Punjab":
                state_matches = "punjab" in admin1

            else:
                state_matches = "himachal" in admin1

            if state_matches:
                location = item
                break
                # STEP 4: LOCATION FALLBACK
        if location is None:

            print("District location not found:", location_query)

            # Try searching for the district headquarters/town
            fallback_places = {
                "samba": "Samba",
                "rajouri": "Rajouri",
                "jammu": "Jammu",
                "kathua": "Kathua",
                "udhampur": "Udhampur",
                "reasi": "Reasi",
                "poonch": "Poonch",
                "doda": "Doda",
                "kishtwar": "Kishtwar",
                "ramban": "Ramban"
            }

            fallback_name = fallback_places.get(district.strip().lower())

            if fallback_name:
                fallback_response = requests.get(
                    "https://geocoding-api.open-meteo.com/v1/search",
                    params={
                        "name": fallback_name,
                        "count": 10,
                        "language": "en",
                        "format": "json",
                        "countryCode": "IN"
                    },
                    timeout=10
                )

                fallback_response.raise_for_status()

                fallback_results = (
                    fallback_response.json().get("results", [])
                )

                for item in fallback_results:
                    country = item.get("country_code", "").upper()
                    place_name = item.get("name", "").strip().lower()

                    if country == "IN" and place_name == fallback_name.lower():
                        location = item
                        break

            if location is None:
                return {
                    "success": False,
                    "message": (
                        f"Could not find a weather location for "
                        f"{location_query}. Please try again later."
                    )
                }, 404
            
        # STEP 5: GET LIVE WEATHER
        weather_response = requests.get(
            "https://api.open-meteo.com/v1/forecast",
            params={
                "latitude": location["latitude"],
                "longitude": location["longitude"],
                "current": (
                    "temperature_2m,"
                    "relative_humidity_2m,"
                    "precipitation,"
                    "wind_speed_10m"
                ),
                "daily": (
                    "temperature_2m_max,"
                    "temperature_2m_min,"
                    "precipitation_probability_max,"
                    "precipitation_sum"
                ),
                "forecast_days": 7,
                "timezone": "auto"
            },
            timeout=10
        )

        weather_response.raise_for_status()

        weather_json = weather_response.json()

        current = weather_json.get("current", {})
        daily = weather_json.get("daily", {})

        # STEP 6: Get sowing recommendation
        sowing_info = get_sowing_recommendation(
            crop=crop,
            state=state_name,
            district=district
        )

        print("SOWING API RESPONSE:", sowing_info)

        # STEP 7: Return weather and sowing information
        return {
            "success": True,

            "crop": crop.title(),
            "state": state_name,
            "district": location.get("name", district),
            "checked_on": date.today().isoformat(),

            "weather": {
                "temperature_c": current.get("temperature_2m"),

                "humidity_percent": current.get(
                    "relative_humidity_2m"
                ),

                "precipitation_mm": current.get("precipitation"),

                "wind_speed_kmh": current.get("wind_speed_10m"),

                "forecast_dates": daily.get("time", []),

                "max_temperature_c": daily.get(
                    "temperature_2m_max", []
                ),

                "min_temperature_c": daily.get(
                    "temperature_2m_min", []
                ),

                "rain_probability_percent": daily.get(
                    "precipitation_probability_max", []
                ),

                "rainfall_mm": daily.get(
                    "precipitation_sum", []
                )
            },

            # Use recommendation from the database
            "sowing": sowing_info
        }

    except requests.RequestException as e:
        print("Sowing API request error:", e)

        return {
            "success": False,
            "message": (
                "Weather or location service is unavailable. "
                "Please check your internet connection and try again."
            )
        }, 502

    except (ValueError, KeyError, TypeError) as e:
        print("Sowing API data error:", e)

        return {
            "success": False,
            "message": "Could not process location or weather data."
        }, 500
# ---------------- MANDI PRICES ----------------

MANDI_API_URL = (
    "https://api.data.gov.in/resource/"
    "9ef84268-d588-465a-a308-a864a43d0070"
)


# ---------------- MANDI API HELPER ----------------

def fetch_mandi_data(params):
    """Fetch mandi data using curl on Windows and requests elsewhere."""

    # Windows par curl.exe use karein
    if os.name == "nt":
        command = [
            "curl.exe",
            "-sS",
            "-G",
            MANDI_API_URL,
        ]

        # API key aur filters sahit sabhi parameters bhejein
        for key, value in params.items():
            command.extend([
                "--data-urlencode",
                f"{key}={value}",
            ])

        command.extend([
            "--connect-timeout", "10",
            "--max-time", "20",
        ])

        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=25,
            check=True,
        )

        result_data = json.loads(result.stdout)

        if result_data.get("error"):
            raise RuntimeError(str(result_data["error"]))

        if result_data.get("status") != "ok":
            raise RuntimeError("Mandi API se valid response nahi mila.")

        return result_data

    # Render aur other platforms par requests use karein
    response = requests.get(
        MANDI_API_URL,
        params=params,
        timeout=(10, 30),
    )

    try:
        response.raise_for_status()
        result_data = response.json()

        if result_data.get("error"):
            raise RuntimeError(str(result_data["error"]))

        if result_data.get("status") != "ok":
            raise RuntimeError("Mandi API se valid response nahi mila.")

        return result_data

    finally:
        response.close()


# ---------------- MANDI ROUTE ----------------

@app.route("/mandi")
def mandi():
    prices = []
    error = None
    data = None

    # Search filters
    state = request.args.get("state", "").strip()
    district = request.args.get("district", "").strip()
    commodity = request.args.get("commodity", "").strip()
    market = request.args.get("market", "").strip()

    # Search button dabane par hi API request chalegi
    searched = request.args.get("search") == "1"

    if searched:

        if not DATA_GOV_API_KEY:
            error = "Mandi API key nahi mili. Configuration check karein."

        else:
            params = {
                "api-key": DATA_GOV_API_KEY,
                "format": "json",
                "offset": 0,
                "limit": 100,
            }

            # Apply filters
            if state:
                params["filters[state]"] = state

            if district:
                params["filters[district]"] = district

            if commodity:
                params["filters[commodity]"] = commodity

            if market:
                params["filters[market]"] = market

            # API request with retry
            for attempt in range(3):
                try:
                    data = fetch_mandi_data(params)
                    break

                except (
                    requests.exceptions.RequestException,
                    subprocess.SubprocessError,
                    ValueError,
                    RuntimeError,
                ) as e:

                    print(
                        f"MANDI API ERROR "
                        f"(attempt {attempt + 1}/3): "
                        f"{type(e).__name__}: {e}"
                    )

                    data = None

                    if attempt < 2:
                        time.sleep(2 ** (attempt + 1))
                    else:
                        error = (
                            "Mandi prices abhi load nahi ho pa rahe hain. "
                            "Kripya kuch der baad dobara try karein."
                        )

            # Convert API records into template data
            if data is not None:

                for item in data.get("records", []):

                    modal_price = item.get("modal_price")

                    try:
                        price_per_kg = (
                            round(float(modal_price) / 100, 2)
                            if modal_price not in ("", None)
                            else None
                        )
                    except (ValueError, TypeError):
                        price_per_kg = None

                    prices.append({
                        "crop": item.get("commodity", "Unknown"),
                        "state": item.get("state", ""),
                        "district": item.get("district", ""),
                        "market": item.get("market", ""),
                        "variety": item.get("variety", ""),
                        "arrival_date": item.get("arrival_date", ""),
                        "min_price": item.get("min_price", ""),
                        "max_price": item.get("max_price", ""),
                        "price": (
                            modal_price
                            if modal_price is not None
                            else ""
                        ),
                        "price_per_kg": price_per_kg,
                        "unit": "per quintal",
                        "source": "data.gov.in / AGMARKNET",
                    })

                if not prices and not error:
                    error = "Is selection ke liye mandi data nahi mila."

    return render_template(
        "mandi.html",
        prices=prices,
        error=error,
        state=state,
        district=district,
        commodity=commodity,
        market=market,
        searched=searched,
    )



# ---------------- KISAN YOJANA ----------------

@app.route("/yojana")
def yojana():
    return render_template("yojana.html")

# ---------------- FARMING CALCULATOR ----------------

@app.route("/calculator", methods=["GET", "POST"])
def calculator():
    result = None
    error = None

    # Conversion factors: 1 unit = hectares
    unit_to_hectare = {
        "hectare": 1.0,
        "acre": 0.4046856422,
        "square_meter": 0.0001,
        "square_feet": 0.000009290304,
        "kanal": 0.0505857,
        "marla": 0.00252929,
    }

    # Approximate regional bigha conversions.
    # These are examples only; local definitions may differ.
    bigha_to_hectare = {
        "uttar_pradesh": 0.25,
        "bihar": 0.25,
        "rajasthan": 0.16,
        "assam": 0.1338,
        "west_bengal": 0.1338,
    }

    if request.method == "POST":
        try:
            area = float(request.form.get("area", 0))
            unit = request.form.get("unit", "hectare")
            state = request.form.get("state", "").lower()
            seed_rate = float(request.form.get("seed_rate", 0))
            cost_per_kg = float(request.form.get("cost_per_kg", 0))

            if area <= 0:
                error = "Farm area must be greater than zero."

            elif seed_rate < 0 or cost_per_kg < 0:
                error = "Seed rate and seed price cannot be negative."

            elif unit == "bigha" and state not in bigha_to_hectare:
                error = (
                    "Please select a supported state for Bigha conversion, "
                    "or choose Acre or Hectare."
                )

            elif unit not in unit_to_hectare and unit != "bigha":
                error = "Please select a valid land unit."

            else:
                if unit == "bigha":
                    conversion_factor = bigha_to_hectare[state]
                else:
                    conversion_factor = unit_to_hectare[unit]

                area_in_hectares = area * conversion_factor
                seed_quantity = area_in_hectares * seed_rate
                total_cost = seed_quantity * cost_per_kg

                result = {
                    "area_in_hectares": round(area_in_hectares, 4),
                    "seed_quantity": round(seed_quantity, 2),
                    "total_cost": round(total_cost, 2),
                    "unit": unit,
                    "area": area,
                }

        except (ValueError, TypeError):
            error = "Please enter valid numeric values."

    return render_template(
        "calculator.html",
        result=result,
        error=error
    )

# ================= ADMIN LOGIN =================

@app.route("/admin-login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")

        print("Username configured:", bool(ADMIN_USERNAME))
        print("Password configured:", bool(ADMIN_PASSWORD))

        if (
            ADMIN_USERNAME
            and ADMIN_PASSWORD
            and username == ADMIN_USERNAME
            and password == ADMIN_PASSWORD
        ):
            session.clear()
            session["admin_logged_in"] = True

            return redirect(url_for("admin_requests"))

        return render_template(
            "admin_login.html",
            error="Invalid username or password."
        ), 401

    return render_template("admin_login.html")


def admin_required(view_function):
    @wraps(view_function)
    def wrapped_view(*args, **kwargs):
        if not session.get("admin_logged_in"):
            return redirect(url_for("admin_login"))

        return view_function(*args, **kwargs)

    return wrapped_view
# ================= ADMIN LOGOUT =================

@app.route("/admin-logout")
def admin_logout():
    session.clear()
    return redirect(url_for("admin_login"))

@app.route("/admin/requests")
@admin_required
def admin_requests():
    conn = get_db()

    try:
        requests_list = conn.execute("""
            SELECT *
            FROM farmer_requests
            ORDER BY created_at DESC
        """).fetchall()

        total = len(requests_list)

        pending = sum(
            1 for item in requests_list
            if item["status"] == "Pending"
        )

        in_progress = sum(
            1 for item in requests_list
            if item["status"] == "In Progress"
        )

        resolved = sum(
            1 for item in requests_list
            if item["status"] == "Resolved"
        )

        return render_template(
            "admin_requests.html",
            requests=requests_list,
            total=total,
            pending=pending,
            in_progress=in_progress,
            resolved=resolved
        )

    finally:
        conn.close()

# ================= REQUEST DETAIL =================

@app.route("/admin/request/<int:request_id>", methods=["GET", "POST"])
@admin_required
def admin_request_detail(request_id):
    conn = get_db()

    try:
        item = conn.execute(
            "SELECT * FROM farmer_requests WHERE id = ?",
            (request_id,)
        ).fetchone()

        if item is None:
            return "Request not found", 404

        if request.method == "POST":
            status = request.form.get("status", "").strip()
            admin_reply = request.form.get("admin_reply", "").strip()

            allowed_statuses = {
                "Pending",
                "In Progress",
                "Resolved"
            }

            if status not in allowed_statuses:
                return "Invalid status", 400

            if len(admin_reply) > 3000:
                return "Reply is too long", 400

            conn.execute("""
                UPDATE farmer_requests
                SET status = ?, admin_reply = ?
                WHERE id = ?
            """, (status, admin_reply, request_id))

            conn.commit()

            return redirect(
                url_for(
                    "admin_request_detail",
                    request_id=request_id
                )
            )

        return render_template(
            "admin_request_detail.html",
            item=item
        )

    finally:
        conn.close()

# ================= FARMER REQUEST TRACKING =================

@app.route("/track-request", methods=["GET", "POST"])
def track_request():
    item = None
    error = None

    if request.method == "POST":
        request_id = request.form.get("request_id", "").strip()

        if not request_id.isdigit():
            error = "Please enter a valid Request ID."
        else:
            conn = get_db()

            try:
                item = conn.execute("""
                    SELECT id, state, district, crop_name,
                           problem_type, description, status,
                           admin_reply, created_at
                    FROM farmer_requests
                    WHERE id = ?
                """, (int(request_id),)).fetchone()

                if item is None:
                    error = "Request not found. Please check your Request ID."

            except sqlite3.Error:
                app.logger.exception("Request tracking database error")
                error = "Unable to check your request right now."

            finally:
                conn.close()

    return render_template(
        "track_request.html",
        item=item,
        error=error
    )
# ================= CROP DISEASE SEED DATA =================


def seed_crop_diseases():
    conn = get_db()

    records = [
        (
            "Rice",
            "Rice Blast",
            "Spindle-shaped leaf spots with a grey centre and dark brown margins.",
            "Use locally recommended resistant varieties and consult an agricultural expert.",
            "https://agritech.tnau.ac.in/crop_protection/rice_diseases/rice_1.html",
            0
        ),
        (
            "Wheat",
            "Yellow Rust",
            "Yellow powdery stripes running parallel to the veins on wheat leaves.",
            "Inspect the crop regularly and consult the local agriculture department.",
            "https://agriculture.hp.gov.in/en/disease-control/",
            0
        ),
        (
            "Maize",
            "Maize Leaf Blight",
            "Elongated yellowish to brown leaf spots on maize leaves.",
            "Inspect plants regularly and seek expert confirmation.",
            "https://www.rkbassam.aau.ac.in/maize",
            0
        ),
        (
            "Tomato",
            "Early Blight",
            "Brown spots with concentric rings on older leaves; affected leaves may turn yellow.",
            "Avoid prolonged leaf wetness, remove infected plant debris safely, and consult a local agricultural expert for suitable management.",
            "https://agritech.tnau.ac.in/",
            0
        ),
        (
            "Mustard",
            "Alternaria Blight",
            "Dark brown to black spots on leaves, often with concentric rings; spots may also occur on stems and pods.",
            "Use healthy seed, follow locally recommended crop practices, and consult an agricultural expert for confirmation and suitable management.",
            "https://agritech.tnau.ac.in/",
            0
        )
    ]

    try:
        for record in records:
            conn.execute("""
                INSERT INTO crop_diseases
                (crop_name, disease_name, symptoms, prevention, source, verified)
                SELECT ?, ?, ?, ?, ?, ?
                WHERE NOT EXISTS (
                    SELECT 1 FROM crop_diseases
                    WHERE LOWER(crop_name) = LOWER(?)
                    AND LOWER(disease_name) = LOWER(?)
                )
            """, record + (record[0], record[1]))

        conn.commit()

    finally:
        conn.close()

        # ---------------- FARMER HELP ----------------

@app.route("/farmer-help", methods=["GET", "POST"])
def farmer_help():

    if request.method == "POST":

        farmer_name = request.form.get(
            "farmer_name", ""
        ).strip()

        mobile = request.form.get(
            "mobile", ""
        ).strip()

        state = request.form.get(
            "state", ""
        ).strip()

        district = request.form.get(
            "district", ""
        ).strip()

        crop_name = request.form.get(
            "crop_name", ""
        ).strip()

        problem_type = request.form.get(
            "problem_type", ""
        ).strip()

        description = request.form.get(
            "description", ""
        ).strip()

        # Validate required fields
        if not all([
            state,
            district,
            crop_name,
            problem_type,
            description
        ]):
            return render_template(
                "farmer_help.html",
                error="Please fill in all required fields."
            ), 400

        if len(description) > 3000:
            return render_template(
                "farmer_help.html",
                error="Description must be 3000 characters or fewer."
            ), 400

        if len(farmer_name) > 100 or len(district) > 100:
            return render_template(
                "farmer_help.html",
                error="Please check the name and district fields."
            ), 400

        if mobile and (
            len(mobile) != 10 or not mobile.isdigit()
        ):
            return render_template(
                "farmer_help.html",
                error="Please enter a valid 10-digit mobile number."
            ), 400

        # ---------------- IMAGE VALIDATION ----------------

        image = request.files.get("problem_image")
        image_filename = None
        image_path = None

        if image and image.filename:

            original_filename = secure_filename(image.filename)

            if (
                not original_filename
                or "." not in original_filename
            ):
                return render_template(
                    "farmer_help.html",
                    error="Please upload a valid image file."
                ), 400

            extension = original_filename.rsplit(
                ".", 1
            )[1].lower()

            if extension not in ALLOWED_EXTENSIONS:
                return render_template(
                    "farmer_help.html",
                    error="Only JPG, JPEG, PNG and WebP images are allowed."
                ), 400

            # Check actual uploaded file size
            image.seek(0, os.SEEK_END)
            image_size = image.tell()
            image.seek(0)

            if image_size == 0:
                return render_template(
                    "farmer_help.html",
                    error="The uploaded image is empty."
                ), 400

            if image_size > MAX_IMAGE_SIZE:
                return render_template(
                    "farmer_help.html",
                    error="Image size must not exceed 5 MB."
                ), 400

            # Generate a unique filename
            image_filename = (
                f"{uuid.uuid4().hex}.{extension}"
            )

            image_path = os.path.join(
                UPLOAD_FOLDER,
                image_filename
            )

        # ---------------- SAVE REQUEST ----------------

        conn = get_db()

        try:
            # Save image first, if provided
            if image and image_filename:
                image.save(image_path)

            cursor = conn.execute("""
                INSERT INTO farmer_requests (
                    farmer_name,
                    mobile,
                    state,
                    district,
                    crop_name,
                    problem_type,
                    description,
                    image_filename
                )
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                farmer_name or None,
                mobile or None,
                state,
                district,
                crop_name,
                problem_type,
                description,
                image_filename
            ))

            request_id = cursor.lastrowid

            conn.commit()

        except Exception:
            conn.rollback()

            # Remove newly saved image if database saving fails
            if image_path and os.path.isfile(image_path):
                try:
                    os.remove(image_path)
                except OSError:
                    app.logger.exception(
                        "Could not remove unused uploaded image"
                    )

            app.logger.exception(
                "Failed to save farmer support request"
            )

            return render_template(
                "farmer_help.html",
                error="Your request could not be saved. Please try again."
            ), 500

        finally:
            conn.close()

        return render_template(
            "farmer_help.html",
            success=True,
            request_id=request_id
        )

    return render_template("farmer_help.html")

        # ---------------- SOWING CALENDAR MIGRATION ----------------

def migrate_sowing_calendar():
    conn = sqlite3.connect(DATABASE)

    try:
        conn.execute("PRAGMA foreign_keys = OFF")

        conn.execute("""
            CREATE TABLE sowing_calendar_new (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                crop_name TEXT NOT NULL,
                state TEXT NOT NULL,
                district TEXT,
                variety TEXT,
                sowing_start TEXT,
                sowing_end TEXT,
                source TEXT,
                verified INTEGER DEFAULT 0,
                recommended_period TEXT
            )
        """)

        conn.execute("""
            INSERT INTO sowing_calendar_new (
                id, crop_name, state, district, variety,
                sowing_start, sowing_end, source,
                verified, recommended_period
            )
            SELECT
                id, crop_name, state, district, variety,
                sowing_start, sowing_end, source,
                verified, recommended_period
            FROM sowing_calendar
        """)

        conn.execute("DROP TABLE sowing_calendar")

        conn.execute("""
            ALTER TABLE sowing_calendar_new
            RENAME TO sowing_calendar
        """)

        # Remove fake placeholder dates from unverified records
        cursor = conn.execute("""
            UPDATE sowing_calendar
            SET sowing_start = NULL,
                sowing_end = NULL
            WHERE verified = 0
              AND recommended_period = 'Not verified'
              AND sowing_start = '2000-01-01'
              AND sowing_end = '2000-01-02'
        """)

        print("Placeholder records cleaned:", cursor.rowcount)

        conn.commit()
        print("Sowing calendar migration completed successfully.")

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()
        
   # ================= SEED SOWING CALENDAR =================

def seed_sowing_calendar():

    conn = get_db()

    districts = [
        "Jammu",
        "Samba",
        "Kathua",
        "Udhampur",
        "Reasi",
        "Rajouri",
        "Poonch",
        "Doda",
        "Kishtwar",
        "Ramban"
    ]

    crops = [
    "Maize",
    "Rice",
    "Wheat",
    "Mustard",
    "Potato",
    "Tomato",
    "Onion",
    "Gram",
    "Peas",
    "Soybean",
    "Bajra"
    ]

    # Dates are intentionally left unverified.
    # Add confirmed dates only after checking official sources.
    records = []

    for district in districts:
        for crop in crops:

            records.append((
    crop,
    "Jammu",
    district,
    "Local variety",
    None,
    None,
    "Official district-wise sowing dates require confirmation",
    0,
    "Not verified"
))

    try:
        for record in records:

            conn.execute("""
                INSERT INTO sowing_calendar (
                    crop_name,
                    state,
                    district,
                    variety,
                    sowing_start,
                    sowing_end,
                    source,
                    verified,
                    recommended_period
                )
                SELECT ?, ?, ?, ?, ?, ?, ?, ?, ?
                WHERE NOT EXISTS (
                    SELECT 1
                    FROM sowing_calendar
                    WHERE LOWER(TRIM(crop_name)) =
                          LOWER(TRIM(?))
                      AND LOWER(TRIM(state)) =
                          LOWER(TRIM(?))
                      AND LOWER(TRIM(COALESCE(district, ''))) =
                          LOWER(TRIM(?))
                )
            """, record + (
                record[0],
                record[1],
                record[2]
            ))

        conn.commit()

        print(
            "Sowing calendar seed completed for "
            "10 districts and 4 crops."
        )

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()


        # ================= UPDATE RAJOURI SOWING DATA =================

def update_rajouri_sowing_data():
    conn = get_db()

    updates = [
        (
            "2026-05-01",
            "2026-06-30",
            "Jammu Agriculture Department - Crop and District Wise Sowing & Harvesting Time",
            1,
            "May–June",
            "Rice",
            "Jammu",
            "Rajouri"
        ),
        (
            "2026-10-01",
            "2026-12-31",
            "Jammu Agriculture Department - Crop and District Wise Sowing & Harvesting Time",
            1,
            "October–December",
            "Wheat",
            "Jammu",
            "Rajouri"
        )
    ]

    try:
        for record in updates:
            conn.execute("""
                UPDATE sowing_calendar
                SET sowing_start = ?,
                    sowing_end = ?,
                    source = ?,
                    verified = ?,
                    recommended_period = ?
                WHERE LOWER(TRIM(crop_name)) = LOWER(?)
                  AND LOWER(TRIM(state)) = LOWER(?)
                  AND LOWER(TRIM(district)) = LOWER(?)
            """, record)

        conn.commit()
        print("Rajouri Rice and Wheat records updated.")

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()

    # -------- OFFICIAL JAMMU DIVISION SOWING CALENDAR --------

def update_jammu_sowing_calendar():

    from datetime import date
    import sqlite3

    districts = [
        "Jammu",
        "Samba",
        "Kathua",
        "Udhampur",
        "Ramban",
        "Reasi",
        "Rajouri",
        "Doda",
        "Kishtwar",
        "Poonch"
    ]

    # Month ranges taken from the Jammu Agriculture
    # Department's district-wise sowing calendar.
    #
    # Format: (start_month, end_month, displayed_period)
    district_periods = {
    "Jammu": {
        "Rice": (6, 7, "June–July"),
        "Maize": (3, 6, "March–June"),
        "Wheat": (10, 12, "October–December"),
        "Mustard": (10, 11, "October–November"),
        "Potato": (2, 3, "February–March"),
        "Tomato": (2, 3, "February–March"),
        "Onion": (11, 1, "November–January"),
        "Gram": (10, 11, "October–November"),
        "Peas": (10, 11, "October–November"),
        "Soybean": (6, 7, "June–July"),
        "Bajra": (6, 7, "June–July")
    },

    "Samba": {
        "Rice": (6, 7, "June–July"),
        "Maize": (3, 6, "March–June"),
        "Wheat": (10, 12, "October–December"),
        "Mustard": (10, 11, "October–November"),
        "Potato": (2, 3, "February–March"),
        "Tomato": (2, 3, "February–March"),
        "Onion": (11, 1, "November–January"),
        "Gram": (10, 11, "October–November"),
        "Peas": (10, 11, "October–November"),
        "Soybean": (6, 7, "June–July"),
        "Bajra": (6, 7, "June–July")
    },

    "Kathua": {
        "Rice": (6, 7, "June–July"),
        "Maize": (3, 6, "March–June"),
        "Wheat": (10, 12, "October–December"),
        "Mustard": (10, 11, "October–November"),
        "Potato": (2, 3, "February–March"),
        "Tomato": (2, 3, "February–March"),
        "Onion": (11, 1, "November–January"),
        "Gram": (10, 11, "October–November"),
        "Peas": (10, 11, "October–November"),
        "Soybean": (6, 7, "June–July"),
        "Bajra": (6, 7, "June–July")
    },

    "Udhampur": {
        "Rice": (6, 7, "June–July"),
        "Maize": (4, 5, "April–May"),
        "Wheat": (10, 12, "October–December"),
        "Mustard": (10, 11, "October–November"),
        "Potato": (2, 3, "February–March"),
        "Tomato": (2, 3, "February–March"),
        "Onion": (11, 1, "November–January"),
        "Gram": (10, 11, "October–November"),
        "Peas": (10, 11, "October–November"),
        "Soybean": (6, 7, "June–July"),
        "Bajra": (6, 7, "June–July")
    },

    "Ramban": {
        "Rice": (5, 6, "May–June"),
        "Maize": (4, 5, "April–May"),
        "Wheat": (10, 12, "October–December"),
        "Mustard": (10, 11, "October–November"),
        "Potato": (3, 4, "March–April"),
        "Tomato": (3, 4, "March–April"),
        "Onion": (10, 11, "October–November"),
        "Gram": (10, 11, "October–November"),
        "Peas": (10, 11, "October–November"),
        "Soybean": (6, 7, "June–July"),
        "Bajra": (6, 7, "June–July")
    },

    "Reasi": {
        "Rice": (6, 7, "June–July"),
        "Maize": (4, 5, "April–May"),
        "Wheat": (10, 12, "October–December"),
        "Mustard": (10, 11, "October–November"),
        "Potato": (2, 3, "February–March"),
        "Tomato": (2, 3, "February–March"),
        "Onion": (11, 1, "November–January"),
        "Gram": (10, 11, "October–November"),
        "Peas": (10, 11, "October–November"),
        "Soybean": (6, 7, "June–July"),
        "Bajra": (6, 7, "June–July")
    },

    "Rajouri": {
        "Rice": (5, 6, "May–June"),
        "Maize": (4, 5, "April–May"),
        "Wheat": (10, 12, "October–December"),
        "Mustard": (10, 11, "October–November"),
        "Potato": (3, 4, "March–April"),
        "Tomato": (3, 4, "March–April"),
        "Onion": (10, 11, "October–November"),
        "Gram": (10, 11, "October–November"),
        "Peas": (10, 11, "October–November"),
        "Soybean": (6, 7, "June–July"),
        "Bajra": (6, 7, "June–July")
    },

    "Doda": {
        "Rice": (5, 6, "May–June"),
        "Maize": (4, 5, "April–May"),
        "Wheat": (10, 12, "October–December"),
        "Mustard": (10, 11, "October–November"),
        "Potato": (3, 4, "March–April"),
        "Tomato": (3, 4, "March–April"),
        "Onion": (10, 11, "October–November"),
        "Gram": (10, 11, "October–November"),
        "Peas": (10, 11, "October–November"),
        "Soybean": (6, 7, "June–July"),
        "Bajra": (6, 7, "June–July")
    },

    "Kishtwar": {
        "Rice": (5, 6, "May–June"),
        "Maize": (4, 5, "April–May"),
        "Wheat": (10, 12, "October–December"),
        "Mustard": (10, 11, "October–November"),
        "Potato": (3, 4, "March–April"),
        "Tomato": (3, 4, "March–April"),
        "Onion": (10, 11, "October–November"),
        "Gram": (10, 11, "October–November"),
        "Peas": (10, 11, "October–November"),
        "Soybean": (6, 7, "June–July"),
        "Bajra": (6, 7, "June–July")
    },

    "Poonch": {
        "Rice": (5, 6, "May–June"),
        "Maize": (4, 5, "April–May"),
        "Wheat": (10, 12, "October–December"),
        "Mustard": (10, 11, "October–November"),
        "Potato": (3, 4, "March–April"),
        "Tomato": (3, 4, "March–April"),
        "Onion": (10, 11, "October–November"),
        "Gram": (10, 11, "October–November"),
        "Peas": (10, 11, "October–November"),
        "Soybean": (6, 7, "June–July"),
        "Bajra": (6, 7, "June–July")
    }
}

    source = (
        "Jammu Agriculture Department - Crop and District "
        "Wise Sowing & Harvesting Time; month ranges "
        "converted to date boundaries"
    )

    today = date.today()
    year = today.year

    conn = get_db()

    try:
        for district, crops in district_periods.items():

            for crop, period in crops.items():

                start_month, end_month, label = period

                start_year = year
                end_year = year

                # October–January crosses into the next year.
                if start_month > end_month:
                    end_year = year + 1

                start_date = date(
                    start_year, start_month, 1
                )

                # Last day of the ending month.
                if end_month == 12:
                    end_date = date(end_year, 12, 31)
                else:
                    from calendar import monthrange

                    last_day = monthrange(
                        end_year, end_month
                    )[1]

                    end_date = date(
                        end_year, end_month, last_day
                    )

                conn.execute("""
                    UPDATE sowing_calendar
                    SET
                        sowing_start = ?,
                        sowing_end = ?,
                        source = ?,
                        verified = 1,
                        recommended_period = ?
                    WHERE LOWER(TRIM(crop_name)) = LOWER(?)
                      AND LOWER(TRIM(state)) = 'jammu'
                      AND LOWER(TRIM(district)) = LOWER(?)
                """, (
                    start_date.isoformat(),
                    end_date.isoformat(),
                    source,
                    label,
                    crop,
                    district
                ))

                # Insert if the district/crop record is missing.
                exists = conn.execute("""
                    SELECT id
                    FROM sowing_calendar
                    WHERE LOWER(TRIM(crop_name)) = LOWER(?)
                      AND LOWER(TRIM(state)) = 'jammu'
                      AND LOWER(TRIM(district)) = LOWER(?)
                    LIMIT 1
                """, (crop, district)).fetchone()

                if exists is None:
                    conn.execute("""
                        INSERT INTO sowing_calendar (
                            crop_name,
                            state,
                            district,
                            variety,
                            sowing_start,
                            sowing_end,
                            source,
                            verified,
                            recommended_period
                        )
                        VALUES (?, 'Jammu', ?, 'Local variety',
                                ?, ?, ?, 1, ?)
                    """, (
                        crop,
                        district,
                        start_date.isoformat(),
                        end_date.isoformat(),
                        source,
                        label
                    ))

        conn.commit()

        print(
            "Jammu Division sowing calendar updated "
            "for 10 districts and 3 crops."
        )

    except Exception:
        conn.rollback()
        raise

    finally:
        conn.close()
# ---------------- START APPLICATION ----------------

def initialize_application():
    with app.app_context():
        print("Starting database initialization...")

        init_db()
        print("init_db() completed.")

        seed_sowing_calendar()
        print("Sowing calendar initialization completed.")

        seed_crop_diseases()
        print("Crop disease seeding completed.")

        update_rajouri_sowing_data()
        update_jammu_sowing_calendar()

        conn = get_db()

        try:
            tables = conn.execute("""
                SELECT name
                FROM sqlite_master
                WHERE type = 'table'
                  AND name = 'crop_diseases'
            """).fetchone()

            if tables:
                print("SUCCESS: crop_diseases table exists.")
            else:
                print("ERROR: crop_diseases table was not created.")

        finally:
            conn.close()

        print("Application database initialization completed.")


initialize_application()

if __name__ == "__main__":
    app.run(debug=True)

