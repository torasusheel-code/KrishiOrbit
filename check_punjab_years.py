import os
import subprocess
import json
import time
from dotenv import load_dotenv

# Load API key from .env
load_dotenv()

API_KEY = os.getenv("CROP_PRODUCTION_API_KEY")

API_URL = (
    "https://api.data.gov.in/resource/"
    "35be999b-0208-4354-b557-f6ca9a5355de"
)


def get_year_records(year):
    """Fetch Punjab crop production records for a specific year."""

    command = [
        "curl.exe",
        "-sS",
        "--fail",
        "--connect-timeout", "10",
        "--max-time", "40",
        "-G", API_URL,
        "--data-urlencode", f"api-key={API_KEY}",
        "--data-urlencode", "format=json",
        "--data-urlencode", "limit=1",
        "--data-urlencode", "filters[state_name]=Punjab",
        "--data-urlencode", f"filters[crop_year]={year}",
    ]

    result = subprocess.run(
        command,
        capture_output=True,
        text=True,
        timeout=45,
        check=False,
    )

    if result.returncode != 0:
        raise RuntimeError(
            f"API request failed for {year}: {result.stderr[:300]}"
        )

    data = json.loads(result.stdout)

    if data.get("error"):
        raise RuntimeError(data["error"])

    return int(data.get("total", 0))


def main():
    if not API_KEY:
        print("ERROR: CROP_PRODUCTION_API_KEY .env mein nahi mili.")
        return

    available_years = []
    unavailable_years = []
    failed_year = None

    print("\nPunjab Crop Production Year Check")
    print("-" * 40)

    for year in range(1997, 2026):
        # API requests ke beech 3 seconds ka gap
        time.sleep(3)

        print(f"Checking year {year}...", flush=True)

        try:
            total = get_year_records(year)

            if total > 0:
                available_years.append(year)
                print(f"{year}: {total} records - Available")
            else:
                unavailable_years.append(year)
                print(f"{year}: 0 records - Not found")

        except Exception as error:
            failed_year = year
            print(f"{year}: ERROR - {error}")
            print("Checking stopped because the API request failed.")
            break

    print("\n" + "=" * 40)
    print("AVAILABLE PUNJAB YEARS")
    print(available_years)

    print("\nYEARS WITH ZERO RECORDS")
    print(unavailable_years)

    if failed_year is not None:
        print(f"\nChecking incomplete. Failed at year: {failed_year}")
        print("Years after this year have not been checked.")
    else:
        print("\nAll requested years were checked successfully.")

    print("\nNote: Zero records means the API returned no")
    print("records for this state-year filter.")


if __name__ == "__main__":
    main()
