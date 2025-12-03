<div align="center">
  <img width="400" alt="Logo" src="pictures/ExcelIMDb.png"></a>
  <br>
  <h1>Excel x IMDb</h1>
  This project is a used to display various statistics regarding watched movies in real life. If the user tracks when and how often they watched a movie in an Excel sheet, this tool can enrich the data with information from the IMDb API like ratings, genres, cast, and more and display it in a GUI for "interesting" visualization.

  ---

  <!-- Placeholder for badges -->
  ![GitHub License](https://img.shields.io/github/license/casudo/life-stats-movies) ![GitHub release (with filter)](https://img.shields.io/github/v/release/casudo/life-stats-movies)
</div>

> [!NOTE]
> This is a hobby project. Feel free to create issues and contribute.

##### Table of Contents

- [Prerequisites](#prerequisites)
- [Usage](#usage)
- [Step 0: Dependencies](#step-0-dependencies)
- [Step 1: Preparation](#step-1-preparation)
- [Step 2: Run the script](#step-2-run-the-script)
- [Planned for the future](#planned-for-the-future)
  - [Thanks to](#thanks-to)
  - [License](#license)

---

## Prerequisites

You need to be a movie nerd who tracks their watched movies in an Excel sheet... if you haven't passed out from boredom yet, this tool might be for you.

If you keep track of the movies you've watched in this fashion..

| Title           | Last Seen  | Rating | Notes         |
| --------------- | ---------- | ------ | ------------- |
| Inception       | 2025-03-19 | 9/10   | Mind-blowing  |
| Matrix Reloaded | 2025-07-15 | 6/10   | Part 2        |
| The Old Guard 2 | 2025-08-10 | 3/10   | Waste of time |

..you can use this tool to enrich your data with IMDb information such as release years, actors, genres, and ratings. You can then use this enriched data to visualize your movie-watching habits in interesting ways.

## Usage

## Step 0: Dependencies

Make sure you have all dependencies installed to run the script.

## Step 1: Preparation

Copy the `env.example` file to `.env` and adjust the constants as needed. Make sure that the sheets / rows / columns exist in your Excel file. The list of movies you've watched can be sorted but it's not mandatory.

## Step 2: Run the script

Run the Python script. It will read your Excel file, fetch data from the IMDb API, and enrich your sheet with additional information.
In the end your Excel sheet will look like this:

![Backend Sheet](/pictures/excel_backend_sheet.png)

In there you can freely sort and filter your data.

If you add new movies to your list, simply rerun the script to update your sheet with the latest IMDb data. The script will only fetch data for movies that don't have an IMDb ID yet, ensuring that your existing data remains intact.

---

## Planned for the future

- Little GUI with input for the constants and progress bar
- Support for other languages titles than English
- Dont hardcode the column indexes
- Add way for visualization

---

### Thanks to

- The devs of [imdabapi.dev](https://imdbapi.dev/) for the free IMDb API

---

### License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
