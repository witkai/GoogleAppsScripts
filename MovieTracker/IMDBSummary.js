var released,
    runtime,
    posterUrl,
    imdbTitle,
    year,
    rated,
    actors,
    votes,
    plot,
    writer,
    rating,
    genre,
    imdbId,
    director;

/**
 * Add the IMDB menu
 */
function onOpen() { 
  var subMenus = [{name: "Show Summary", functionName: "showImdbSummary"}];
  SpreadsheetApp.getActiveSpreadsheet().addMenu("IMDB", subMenus);
}

/**
 * Show the IMDB summary in a window
 */
function showImdbSummary() {
  var sheet = SpreadsheetApp.getActive().getSheets()[0];
  var movieTitle = SpreadsheetApp.getActiveRange().getValue();
  // get movie from feed
  getMovieFromImdb(movieTitle);
  // UI
  var app = UiApp.createApplication();
  app.setTitle("IMDP Summary");
  app.setHeight(900);
  app.setWidth(900);
  // details
  var grid = app.createGrid(12,2);
  grid.setCellPadding(3);
  var row = 0;
  grid.setWidget(row, 0, app.createLabel("Title:"));
  grid.setWidget(row, 1, app.createLabel(imdbTitle));
  row += 1;
  grid.setWidget(row, 0, app.createLabel("Year:"));
  grid.setWidget(row, 1, app.createLabel(year));
  row += 1;
  grid.setWidget(row, 0, app.createLabel("Released:"));
  grid.setWidget(row, 1, app.createLabel(released));
  row += 1;
  grid.setWidget(row, 0, app.createLabel("Runtime:"));
  grid.setWidget(row, 1, app.createLabel(runtime));
  row += 1;
  grid.setWidget(row, 0, app.createLabel("Rated:"));
  grid.setWidget(row, 1, app.createLabel(rated));
  row += 1;
  grid.setWidget(row, 0, app.createLabel("Actors:"));
  grid.setWidget(row, 1, app.createLabel(actors));
  row += 1;
  grid.setWidget(row, 0, app.createLabel("Genre:"));
  grid.setWidget(row, 1, app.createLabel(genre));
  row += 1;
  grid.setWidget(row, 0, app.createLabel("Rating:"));
  grid.setWidget(row, 1, app.createLabel(rating));
  row += 1;
  grid.setWidget(row, 0, app.createLabel("Votes:"));
  grid.setWidget(row, 1, app.createLabel(votes));
  row += 1;
  grid.setWidget(row, 0, app.createLabel("Plot:"));
  grid.setWidget(row, 1, app.createLabel(plot));
  row += 1;
  grid.setWidget(row, 0, app.createLabel("Director:"));
  grid.setWidget(row, 1, app.createLabel(director));
  row += 1;
  grid.setWidget(row, 0, app.createLabel("Writer:"));
  grid.setWidget(row, 1, app.createLabel(writer));
  // image
  var image = app.createImage();
  image.setUrl(posterUrl);
  // panels
  var hPanel = app.createHorizontalPanel();
  hPanel.add(grid);
  hPanel.add(image);
  var vPanel = app.createVerticalPanel();
  vPanel.add(hPanel);
  app.add(vPanel);
  // buttons
  var closeButton = app.createButton("Close");
  var closeHandler = app.createServerClickHandler("closeApp");
  closeButton.addClickHandler(closeHandler);  
  closeHandler.addCallbackElement(vPanel);
  vPanel.add(closeButton);
  SpreadsheetApp.getActiveSpreadsheet().show(app);
  // Write data to sheet
  updateMovie();
}

/**
 * Get the movie summary from IMDB
 */
function getMovieFromImdb(title) {
  var baseUrl = "http://www.imdbapi.com/?t=";
  var response = UrlFetchApp.fetch(baseUrl + title);
  if (response.getResponseCode() == 200) {
    result = Utilities.jsonParse(response.getContentText());
    released = result.Released;
    runtime = result.Runtime;
    posterUrl = result.Poster;
    imdbTitle = result.Title;
    year = result.Year;
    rated = result.Rated;
    actors = result.Actors;
    votes = result.imdbVotes;
    plot = result.Plot;
    writer = result.Writer;
    rating = result.imdbRating;
    genre = result.Genre;
    imdbId = result.ID;
    director = result.Director;
  }
}

/**
 * Close the UiApp window
 */
function closeApp() {
  Logger.log("closeApp");
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}
 
/**
 * Write the IMDB data for the selected movie to the sheet.
 */
function updateMovie() {
  var sheet = SpreadsheetApp.getActive().getSheets()[0];
  // the row of the current movie
  var row = SpreadsheetApp.getActiveRange().getRow();
  // write data to cells
  updateValue(sheet, row, 8, rating);
  updateValue(sheet, row, 9, released);
  updateValue(sheet, row, 10, runtime);
  updateValue(sheet, row, 11, genre);
  updateValue(sheet, row, 12, year);
  updateValue(sheet, row, 13, rated);
  updateValue(sheet, row, 14, actors);
  updateValue(sheet, row, 15, votes);
  updateValue(sheet, row, 16, plot);
  updateValue(sheet, row, 17, writer);
  updateValue(sheet, row, 18, director);
  updateValue(sheet, row, 19, posterUrl);
  updateValue(sheet, row, 20, imdbId);
}   

/**
 * Updates a cell in a movie row.
 * sheet, The active sheet that contains the movies
 * row, The row of the movie
 * column, The column for the value (A=1,B=2,C=3,etc.)
 * value, The value to write into the cell
 */
function updateValue(sheet, row, column, value) {
  sheet.getDataRange().getCell(row, column).setValue(value);
}
