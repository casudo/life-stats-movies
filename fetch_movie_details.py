import requests
import openpyxl
from urllib.parse import quote
from datetime import datetime
from os import getenv
from dotenv import load_dotenv

### Load environment variables from .env file
load_dotenv()

### =================================================================================================

EXCEL_PATH = getenv("EXCEL_PATH")

SOURCE_SHEET_NAME = getenv("SOURCE_SHEET_NAME")  # Sheet where users enter movie names
BACKEND_SHEET_NAME = getenv("BACKEND_SHEET_NAME")  # Sheet where IMDb data is stored

SOURCE_MOVIE_COLUMN = getenv("SOURCE_MOVIE_COLUMN")  # Column with movie names in source sheet
SOURCE_STARTING_ROW = int(getenv("SOURCE_STARTING_ROW"))  # First data row in source sheet

BACKEND_MOVIE_COLUMN = getenv("BACKEND_MOVIE_COLUMN")  # Column with movie names in backend sheet
BACKEND_STARTING_ROW = int(getenv("BACKEND_STARTING_ROW"))  # First data row in backend sheet

TIMESTAMP_CELL = getenv("TIMESTAMP_CELL")  # Cell to write the last update timestamp

### =================================================================================================

__version__ = "v1.0.0"
__author__ = "casudo"

def search_movie(title: str) -> dict:
    """Search for a movie by title using IMDb API"""
    url = "https://api.imdbapi.dev/search/titles"
    query = f"?query={quote(title)}&limit=5"
    
    try:
        response = requests.get(url + query)
        if response.status_code == 200:
            return response.json()
        return {}
    except Exception as e:
        print(f"Unexpected error fetching data for '{title}': {e}")
        return {}


def find_imdb_match(title: str) -> str | None:
    """Find IMDb ID for a movie title"""
    search_results = search_movie(title)
    
    if not search_results or "titles" not in search_results:
        return None
    
    title_lower = title.lower()
    
    ### Check each result for a match
    for movie in search_results["titles"]:
        ### Check if it's a movie type
        if movie.get("type") != "movie":
            continue

        ### Check if title matches (either primaryTitle or originalTitle)
        primary_title = movie.get("primaryTitle", "").lower()
        original_title = movie.get("originalTitle", "").lower()
        
        ### INFO: Sadly no aka titles are provided by the API if searched with title

        if title_lower == primary_title or title_lower == original_title:
            return movie.get("id")
    
    return None


def get_movie_details(imdb_id: str) -> dict:
    """Fetch detailed movie information from IMDb API using IMDb ID"""
    url = f"https://api.imdbapi.dev/titles/{imdb_id}"
    
    try:
        response = requests.get(url)
        if response.status_code == 200:
            return response.json()
        return {}
    except Exception as e:
        print(f"Unexpected error fetching details for IMDb ID '{imdb_id}': {e}")
        return {}


def format_runtime(seconds: int) -> str:
    """Convert runtime from seconds to hours and minutes format"""
    if not seconds:
        return "N/A"
    
    hours = seconds // 3600
    minutes = (seconds % 3600) // 60
    
    if hours > 0:
        return f"{hours}h {minutes}min"
    else:
        return f"{minutes}min"


def main(file_path: str) -> None:
    """Sync movies from source sheet to backend, fetch IMDb IDs and enrich data in one go"""
    try:
        ### Load the workbook
        workbook = openpyxl.load_workbook(file_path)
        source_sheet = workbook[SOURCE_SHEET_NAME]
        backend_sheet = workbook[BACKEND_SHEET_NAME]
        
        ### Write timestamp
        current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        backend_sheet[TIMESTAMP_CELL] = f"Last updated: {current_datetime}"
        
        ### Step 1: Read all movie names from source sheet
        print("Reading movies from source sheet...\n")
        source_movies = []
        row = SOURCE_STARTING_ROW
        while True:
            cell_value = source_sheet[f"{SOURCE_MOVIE_COLUMN}{row}"].value
            if not cell_value:
                break
            movie_title = str(cell_value).strip()
            if movie_title:
                source_movies.append(movie_title)
            row += 1
        
        print(f"Found {len(source_movies)} movies in '{SOURCE_SHEET_NAME}' sheet.\n")
        
        ### Step 2: Read existing movies from backend sheet
        existing_movies = {}
        row = BACKEND_STARTING_ROW
        while True:
            movie_name = backend_sheet[f"{BACKEND_MOVIE_COLUMN}{row}"].value
            imdb_id = backend_sheet[f"C{row}"].value
            release_year = backend_sheet[f"D{row}"].value  # Check if enriched
            
            if not movie_name:
                break
            
            movie_title = str(movie_name).strip()
            if movie_title:
                ### Store row number, IMDb ID status, and enrichment status
                has_valid_id = imdb_id and imdb_id != "N/A" and isinstance(imdb_id, str) and imdb_id.startswith("tt")
                is_enriched = release_year is not None and release_year != "" and release_year != "N/A"
                existing_movies[movie_title.lower()] = {
                    "row": row,
                    "has_id": has_valid_id,
                    "imdb_id": imdb_id if has_valid_id else None,
                    "is_enriched": is_enriched
                }
            row += 1
        
        backend_next_row = row  # Next available row in backend
        print(f"Found {len(existing_movies)} existing movies in '{BACKEND_SHEET_NAME}' sheet.\n")
        
        ### Step 3: Process movies
        print("Processing and enriching movies...\n")
        print(f"{'Movie Title':<45} {'IMDb ID':<15} {'Status'}")
        print("-" * 95)
        
        new_count = 0
        enriched_count = 0
        skipped_count = 0
        failed_count = 0
        
        for movie_title in source_movies:
            movie_lower = movie_title.lower()
            
            ### Check if movie already exists in backend with valid IMDb ID and is enriched
            if movie_lower in existing_movies:
                existing = existing_movies[movie_lower]
                if existing["has_id"] and existing["is_enriched"]:
                    # Already has ID and enriched data - skip completely
                    skipped_count += 1
                    continue
                elif existing["has_id"] and not existing["is_enriched"]:
                    # Has IMDb ID but not enriched - use existing ID
                    backend_row = existing["row"]
                    imdb_id = existing["imdb_id"]
                    is_new = False
                    needs_id_lookup = False
                else:
                    # Movie exists but no valid ID - update existing row
                    backend_row = existing["row"]
                    is_new = False
                    needs_id_lookup = True
            else:
                ### New movie - use next available row
                backend_row = backend_next_row
                backend_next_row += 1
                is_new = True
                needs_id_lookup = True
            
            ### Fetch IMDb ID if needed
            if needs_id_lookup:
                imdb_id = find_imdb_match(movie_title)
            else:
                # Already have IMDb ID from existing data - just need enrichment
                pass
            
            if not imdb_id:
                ### No IMDb ID found
                if is_new:
                    backend_sheet[f"{BACKEND_MOVIE_COLUMN}{backend_row}"] = movie_title
                backend_sheet[f"C{backend_row}"] = "N/A"
                print(f"{movie_title:<45} {'N/A':<15} ✗ Not found")
                failed_count += 1
                if is_new:
                    new_count += 1
                continue
            
            ### Fetch detailed movie data
            movie_data = get_movie_details(imdb_id)
            
            if not movie_data or "id" not in movie_data:
                ### API call failed
                if is_new:
                    backend_sheet[f"{BACKEND_MOVIE_COLUMN}{backend_row}"] = movie_title
                    backend_sheet[f"C{backend_row}"] = imdb_id
                    new_count += 1
                print(f"{movie_title:<45} {imdb_id:<15} ⚠ ID found, enrichment failed")
                failed_count += 1
                continue
            
            ### Write movie name and IMDb ID (if new movie or if ID was just fetched)
            if is_new:
                backend_sheet[f"{BACKEND_MOVIE_COLUMN}{backend_row}"] = movie_title
            if needs_id_lookup:
                backend_sheet[f"C{backend_row}"] = imdb_id
            
            ### Enrich with detailed data
            # D: Release Year
            release_year = movie_data.get("startYear", "N/A")
            backend_sheet[f"D{backend_row}"] = release_year
            
            # E: Length (runtime)
            runtime_seconds = movie_data.get("runtimeSeconds", 0)
            length = format_runtime(runtime_seconds)
            backend_sheet[f"E{backend_row}"] = length
            
            # F: Director (Reggisseur)
            directors = movie_data.get("directors", [])
            director_names = ", ".join([d.get("displayName", "") for d in directors]) if directors else "N/A"
            backend_sheet[f"F{backend_row}"] = director_names
            
            # G: Genres (comma separated)
            genres = movie_data.get("genres", [])
            genres_str = ", ".join(genres) if genres else "N/A"
            backend_sheet[f"G{backend_row}"] = genres_str
            
            # H: IMDb Rating
            rating_data = movie_data.get("rating", {})
            imdb_rating = rating_data.get("aggregateRating", "N/A")
            backend_sheet[f"H{backend_row}"] = imdb_rating
            
            # I: Metacritic Rating
            metacritic_data = movie_data.get("metacritic", {})
            metacritic_score = metacritic_data.get("score", "N/A")
            backend_sheet[f"I{backend_row}"] = metacritic_score
            
            # J: Stars / Actors (comma separated)
            stars = movie_data.get("stars", [])
            stars_names = ", ".join([s.get("displayName", "") for s in stars]) if stars else "N/A"
            backend_sheet[f"J{backend_row}"] = stars_names
            
            print(f"{movie_title:<45} {imdb_id:<15} ✓ Enriched")
            enriched_count += 1
            if is_new:
                new_count += 1
        
        ### Save the workbook
        workbook.save(file_path)
        workbook.close()
        
        print(f"\n{'='*95}")
        print("Processing complete!")
        print(f"  New movies added and enriched: {new_count}")
        print(f"  Total enriched this run: {enriched_count}")
        print(f"  Already had IMDb data (skipped): {skipped_count}")
        print(f"  Failed to find IMDb ID: {failed_count}")
        print(f"  Total in source: {len(source_movies)}")
        print(f"  Total in backend: {len(existing_movies) + new_count}")
        print(f"\nTimestamp added to cell {TIMESTAMP_CELL}: {current_datetime}")
        
    except FileNotFoundError:
        print(f"Error: File '{file_path}' not found!")
    except KeyError as e:
        print(f"Error: Sheet not found in the workbook! {e}")
    except PermissionError:
        print(f"Error: Permission denied when trying to save the file '{file_path}'. Please close the file if it's open in another program.")
    except Exception as e:
        print(f"Unexpected error processing Excel file: {e}")


if __name__ == "__main__":
    main(EXCEL_PATH)
