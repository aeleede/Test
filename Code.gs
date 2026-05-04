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

  let updated = 0;
  let alreadyComplete = 0;
  const notFound = [];

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

    const movieData = fetchFromOmdb_(title, year);

    if (!movieData) {
      notFound.push(`"${title}"${year ? ' (' + year + ')' : ''}`);
      continue;
    }

    // Sheet rows are 1-based; data starts at row 2, so sheetRow = i + 2
    const sheetRow = i + 2;

    if (missingMpaa && movieData.Rated && movieData.Rated !== 'N/A') {
      sheet.getRange(sheetRow, 3).setValue(movieData.Rated);
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

    updated++;
    Utilities.sleep(250); // Stay well within OMDB rate limits
  }

  let summary = `Done!\n\n✓ Updated: ${updated} movie(s)\n✓ Already complete: ${alreadyComplete}`;
  if (notFound.length > 0) {
    summary += `\n\nCould not find (${notFound.length}):\n${notFound.join('\n')}\n\nTip: Check that the title in column A matches the OMDB title exactly, and that column H has the correct release year.`;
  }
  SpreadsheetApp.getUi().alert(summary);
}

function fetchFromOmdb_(title, year) {
  const params = [
    't=' + encodeURIComponent(title),
    year ? 'y=' + encodeURIComponent(year) : '',
    'apikey=' + OMDB_API_KEY
  ].filter(Boolean).join('&');

  const url = 'https://www.omdbapi.com/?' + params;

  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const data = JSON.parse(response.getContentText());
    return data.Response === 'True' ? data : null;
  } catch (e) {
    return null;
  }
}
