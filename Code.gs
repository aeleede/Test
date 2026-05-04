// Movie Metadata Auto-Filler
// ============================================================
// SETUP INSTRUCTIONS:
// 1. Get a free OMDB API key at https://www.omdbapi.com/apikey.aspx
// 2. Paste your key in the OMDB_API_KEY constant below
// 3. In your Google Sheet, open Extensions > Apps Script
// 4. Paste this entire file into the editor and click Save
// 5. Reload your spreadsheet — a "Movies" menu will appear
// 6. Click Movies > Fill Missing Metadata to run
//
// COLUMN LAYOUT EXPECTED:
//   A = Movie Title      B = Date Watched    C = MPAA Rating (filled)
//   D = Genre (filled)   E = Your Rating     F = Actors (filled)
//   G = Director (filled) H = Release Year   I = Notes
//
// Only blank cells in C, D, F, G are written — your data is never overwritten.
// If column H is blank, OMDB is searched by title; if multiple movies match
// you will be prompted to pick the correct one, and the year is written back.
// Column C dropdown values are read automatically and matched against OMDB's
// rating string so the selected value is always a valid dropdown option.
// ============================================================

const OMDB_API_KEY = 'YOUR_API_KEY_HERE';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Movies')
    .addItem('Fill Missing Metadata', 'fillMovieMetadata')
    .addToUi();
}

function fillMovieMetadata() {
  if (OMDB_API_KEY === 'YOUR_API_KEY_HERE') {
    SpreadsheetApp.getUi().alert(
      'Setup required\n\nPlease replace YOUR_API_KEY_HERE in the script with your OMDB API key.\nGet a free key at https://www.omdbapi.com/apikey.aspx'
    );
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('No movie rows found (sheet appears empty below row 1).');
    return;
  }

  // Read all rows at once to minimise Sheets API calls
  const allData = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  // 0-based column indices within allData rows:
  // A=0  B=1  C=2  D=3  E=4  F=5  G=6  H=7  I=8
  const COL_TITLE    = 0;
  const COL_MPAA     = 2;
  const COL_GENRE    = 3;
  const COL_ACTORS   = 5;
  const COL_DIRECTOR = 6;
  const COL_YEAR     = 7;

  // Read the allowed MPAA values from column C's dropdown (if one exists)
  const mpaaOptions = getDropdownOptions_(sheet, 2, 3);

  let updated = 0;
  let alreadyComplete = 0;
  const notFound = [];
  const skipped = [];

  for (let i = 0; i < allData.length; i++) {
    const row = allData[i];
    const title    = row[COL_TITLE];
    const mpaa     = row[COL_MPAA];
    const genre    = row[COL_GENRE];
    const actors   = row[COL_ACTORS];
    const director = row[COL_DIRECTOR];
    const year     = row[COL_YEAR];

    if (!title) continue;

    const missingMpaa     = !mpaa;
    const missingGenre    = !genre;
    const missingActors   = !actors;
    const missingDirector = !director;

    if (!missingMpaa && !missingGenre && !missingActors && !missingDirector) {
      alreadyComplete++;
      continue;
    }

    const sheetRow = i + 2; // sheet rows are 1-based, data starts at row 2
    let movieData = null;

    if (year) {
      // Year known — exact lookup
      movieData = fetchByTitle_(title, year);
      if (!movieData) notFound.push(`"${title}" (${year})`);
    } else {
      // No year — search by title and resolve ambiguity
      const result = resolveBySearch_(title, sheetRow, sheet);
      if (result === 'skipped') {
        skipped.push(`"${title}"`);
        continue;
      } else if (result === null) {
        notFound.push(`"${title}"`);
        continue;
      }
      movieData = result;
    }

    if (!movieData) continue;

    if (missingMpaa && movieData.Rated && movieData.Rated !== 'N/A') {
      const mpaaValue = matchDropdownOption_(movieData.Rated, mpaaOptions);
      if (mpaaValue) sheet.getRange(sheetRow, 3).setValue(mpaaValue);
    }
    if (missingGenre && movieData.Genre && movieData.Genre !== 'N/A') {
      sheet.getRange(sheetRow, 4).setValue(movieData.Genre);
    }
    if (missingActors && movieData.Actors && movieData.Actors !== 'N/A') {
      sheet.getRange(sheetRow, 6).setValue(movieData.Actors);
    }
    if (missingDirector && movieData.Director && movieData.Director !== 'N/A') {
      sheet.getRange(sheetRow, 7).setValue(movieData.Director);
    }

    // Write back the year if it was blank and we found one
    if (!year && movieData.Year && movieData.Year !== 'N/A') {
      sheet.getRange(sheetRow, 8).setValue(movieData.Year);
    }

    updated++;
    Utilities.sleep(250); // Stay well within OMDB rate limits
  }

  let summary = `Done!\n\n✓ Updated: ${updated} movie(s)\n✓ Already complete: ${alreadyComplete}`;
  if (skipped.length > 0) {
    summary += `\n\nSkipped (you chose to skip):\n${skipped.join('\n')}`;
  }
  if (notFound.length > 0) {
    summary += `\n\nCould not find (${notFound.length}):\n${notFound.join('\n')}\n\nTip: Check that the title in column A matches the OMDB title exactly.`;
  }
  SpreadsheetApp.getUi().alert(summary);
}

// Searches OMDB by title, handles disambiguation, returns full movie data or
// 'skipped' (user chose to skip) or null (not found / cancelled).
function resolveBySearch_(title, sheetRow, sheet) {
  const results = searchOmdb_(title);

  if (!results || results.length === 0) return null;

  // Only one match — use it without prompting
  if (results.length === 1) {
    return fetchById_(results[0].imdbID);
  }

  // Multiple matches — ask the user to pick
  const ui = SpreadsheetApp.getUi();
  const lines = results.map((m, idx) => `${idx + 1}. ${m.Title} (${m.Year})`);
  const prompt =
    `Multiple matches found for "${title}".\n\n` +
    `${lines.join('\n')}\n\n` +
    `Type the number of the correct movie, or 0 to skip:`;

  const response = ui.prompt('Select Movie', prompt, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return 'skipped';

  const choice = parseInt(response.getResponseText().trim(), 10);

  if (choice === 0) return 'skipped';
  if (isNaN(choice) || choice < 1 || choice > results.length) {
    ui.alert(`Invalid selection for "${title}" — skipping.`);
    return 'skipped';
  }

  return fetchById_(results[choice - 1].imdbID);
}

// OMDB title search — returns array of movie results (type=movie only), max 5.
function searchOmdb_(title) {
  const url = `https://www.omdbapi.com/?s=${encodeURIComponent(title)}&type=movie&apikey=${OMDB_API_KEY}`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());
    if (data.Response !== 'True' || !data.Search) return null;
    return data.Search.slice(0, 5); // Cap at 5 to keep prompts readable
  } catch (e) {
    return null;
  }
}

// Fetch full movie details by IMDB ID.
function fetchById_(imdbId) {
  const url = `https://www.omdbapi.com/?i=${encodeURIComponent(imdbId)}&apikey=${OMDB_API_KEY}`;
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());
    return data.Response === 'True' ? data : null;
  } catch (e) {
    return null;
  }
}

// Returns the allowed values from a dropdown validation on a given cell,
// or null if there is no dropdown / the type is unsupported.
function getDropdownOptions_(sheet, row, col) {
  const rule = sheet.getRange(row, col).getDataValidation();
  if (!rule) return null;

  const type = rule.getCriteriaType();
  const values = rule.getCriteriaValues();

  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return values[0]; // already an array of strings
  }
  if (type === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    return values[0].getValues().flat().map(String).filter(v => v.trim() !== '');
  }
  return null;
}

// Tries to match an OMDB rating string to one of the dropdown options.
// Uses case-insensitive exact match first, then common aliases.
// Returns the matched option string, or the raw value if no dropdown exists,
// or null if a dropdown exists but no match could be found.
function matchDropdownOption_(omdbRated, options) {
  if (!options) return omdbRated; // no dropdown — use value as-is

  // Case-insensitive exact match
  const lower = omdbRated.toLowerCase();
  const exact = options.find(o => o.toLowerCase() === lower);
  if (exact) return exact;

  // Common aliases: maps what OMDB might return to likely dropdown labels
  const aliases = {
    'not rated': ['NR', 'Not Rated', 'Unrated', 'UR'],
    'unrated':   ['NR', 'Not Rated', 'Unrated', 'UR'],
    'nr':        ['NR', 'Not Rated'],
    'ur':        ['UR', 'Unrated', 'NR', 'Not Rated'],
    'g':         ['G'],
    'pg':        ['PG'],
    'pg-13':     ['PG-13'],
    'r':         ['R'],
    'nc-17':     ['NC-17'],
    'tv-g':      ['TV-G', 'G'],
    'tv-pg':     ['TV-PG', 'PG'],
    'tv-14':     ['TV-14', 'PG-13'],
    'tv-ma':     ['TV-MA', 'R'],
  };

  const candidates = aliases[lower] || [];
  for (const candidate of candidates) {
    const match = options.find(o => o.toLowerCase() === candidate.toLowerCase());
    if (match) return match;
  }

  return null; // dropdown exists but no match — leave cell blank rather than set invalid value
}

// Exact title + optional year lookup.
function fetchByTitle_(title, year) {
  const params = [
    't=' + encodeURIComponent(title),
    year ? 'y=' + encodeURIComponent(year) : '',
    'apikey=' + OMDB_API_KEY
  ].filter(Boolean).join('&');

  try {
    const response = UrlFetchApp.fetch(`https://www.omdbapi.com/?${params}`, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());
    return data.Response === 'True' ? data : null;
  } catch (e) {
    return null;
  }
}
