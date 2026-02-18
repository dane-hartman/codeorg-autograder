/**
 * Game Lab Autograder (Google Sheets + Apps Script)
 * Fresh install script: includes setupSheets(), robust structured-output fallback,
 * diagnostics, grading, Code.org fetch, and optional emailer.
 *
 * Default LLM provider is Gemini.
 * Set GEMINI_API_KEY in Project settings → Script properties.
 * Optional: set LLM_PROVIDER=openai and OPENAI_API_KEY.
 */

/** === CONFIG === */
var SHEET_SUB     = 'Submissions';
var SHEET_LEVELS  = 'Levels';function gradeRows_(rowNums) {
  if (!rowNums || !rowNums.length) return;

  var sh = getSheet_(SHEET_SUB);
  var data = sh.getDataRange().getValues();
  var head = headers_(data[0]);
  var critByLevel = loadCriteriaByLevel_();
  var enabledLevels = loadEnabledLevels_();

  rowNums.forEach(function(rowNum) {
    try {
      var r = rowNum - 1;
      var url = (data[r][head.ShareURL] || '').toString().trim();
      var levelId = (data[r][head.LevelID] || '').toString().trim();

      if (!url || !levelId) {
        writeRow_(sh, rowNum, head, { Status: 'No URL/LevelID', Score: 0, Notes: '' });
        return;
      }
      if (!enabledLevels[levelId]) {
        writeRow_(sh, rowNum, head, { Status: 'Level disabled/unknown', Score: 0, Notes: '' });
        return;
      }

      var crits = critByLevel[levelId] || [];
      var maxPts = crits.reduce(function(s,c){ return s + (Number(c.Points)||0); }, 0);
      if (!crits.length) {
        writeRow_(sh, rowNum, head, { Status: 'No criteria found', Score: 0, MaxScore: 0, Notes: '' });
        return;
      }

      var channelId = extractChannelId_(url);
      if (!channelId) {
        writeRow_(sh, rowNum, head, {
          ChannelID: '',
          Score: 0,
          MaxScore: maxPts,
          Status: 'Invalid share link (no ChannelID)',
          Notes: 'Expecting a projects/gamelab/<id> share URL'
        });
        return;
      }

      var fetched = fetchGameLabSource_(channelId);
      if (!fetched || !fetched.ok) {
        writeRow_(sh, rowNum, head, {
          ChannelID: channelId,
          Score: 0,
          MaxScore: maxPts,
          Status: 'Invalid share link or unreadable project',
          Notes: (fetched && fetched.msg) ? fetched.msg : 'Fetch failed'
        });
        return;
      }

      // Grade (may hit cache)
      var res = runCriteria_(fetched.src, crits, levelId);
      var patch = {
        ChannelID: channelId,
        Score: res.score,
        MaxScore: res.max,
        Status: 'OK',
        Notes: res.notes.join(' | ')
      };

      // Compare against existing row; if unchanged AND (hit cache), skip writing
      var prev = {
        ChannelID: data[r][head.ChannelID],
        Score:     data[r][head.Score],
        MaxScore:  data[r][head.MaxScore],
        Status:    data[r][head.Status],
        Notes:     data[r][head.Notes]
      };
      var changed = Object.keys(patch).some(function(k){
        return String(prev[k]||'') !== String(patch[k]||'');
      });

      // We can detect cache hit by peeking into runCriteria_ → openaiGrade_ return.
      // Add a tiny signal: if ANY note starts with ✅/❌ (normal), we can't tell.
      // So: we expose a global lastCacheHit flag set by openaiGrade_ (optional),
      // or simpler: treat "no change" as "no write".
      if (!changed) {
        // nothing changed — avoid rewriting cells
        return;
      }

      writeRow_(sh, rowNum, head, patch);

    } catch (e) {
      writeRow_(sh, rowNum, head, { Status: 'Error', Notes: String(e) });
    }
  });
}

var SHEET_CRIT    = 'Criteria';
// Default provider can be overridden via Script properties:
//   LLM_PROVIDER = 'gemini' | 'openai'
// API keys:
//   GEMINI_API_KEY (recommended default)
//   OPENAI_API_KEY (optional)
var DEFAULT_PROVIDER = 'gemini';
var DEFAULT_MODEL_BY_PROVIDER = {
  gemini: 'gemini-1.5-flash',
  openai: 'gpt-4o'
};

function getLLMProvider_() {
  var p = PropertiesService.getScriptProperties().getProperty('LLM_PROVIDER');
  p = (p || DEFAULT_PROVIDER || 'gemini').toString().trim().toLowerCase();
  if (p !== 'openai' && p !== 'gemini') p = DEFAULT_PROVIDER;
  return p;
}

function getDefaultModel_() {
  var p = getLLMProvider_();
  return (DEFAULT_MODEL_BY_PROVIDER[p] || DEFAULT_MODEL_BY_PROVIDER.gemini);
}

/** === ONE-TIME BUILDER === */
/** Creates the three tabs with headers, starter rows, and LevelID dropdown */
function setupSheets() {
  var ss = SpreadsheetApp.getActive();

  // Create or clear a sheet
  function ensureSheet(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.clear();
    return sh;
  }

  // Submissions
  var sub = ensureSheet(SHEET_SUB);
  var subHeaders = [
    'Timestamp','First','Last','Period','LevelID','ShareURL',
    'ChannelID','Score','MaxScore','Status','Notes','Email','EmailedAt'
  ];
  sub.getRange(1,1,1,subHeaders.length).setValues([subHeaders]).setFontWeight('bold');
  sub.setFrozenRows(1);

  // Levels
  var lev = ensureSheet(SHEET_LEVELS);
  var levHeaders = ['LevelID','LevelName','Enabled','Model'];
  lev.getRange(1,1,1,levHeaders.length).setValues([levHeaders]).setFontWeight('bold');
  var rubric = parseCriteriaTableCsv_();
  var levelIds = Object.keys(rubric.byLevel).sort();
  var levelRows = levelIds.map(function(id){
    return [id, '', true, ''];
  });
  if (levelRows.length) lev.getRange(2,1,levelRows.length,4).setValues(levelRows);
  lev.setFrozenRows(1);

  // Criteria (LLM-only)
  var crit = ensureSheet(SHEET_CRIT);
  // Keep Teacher Notes as a visible column, even though grading currently ignores it.
  var critHeaders = ['LevelID','CriterionID','Points','Type','Description','Notes','Teacher Notes'];
  crit.getRange(1,1,1,critHeaders.length).setValues([critHeaders]).setFontWeight('bold');
  crit.setFrozenRows(1);
  if (rubric.rows.length) crit.getRange(2,1,rubric.rows.length,critHeaders.length).setValues(rubric.rows);

  // LevelID dropdown in Submissions!E:E
  var rule = SpreadsheetApp.newDataValidation()
    .requireFormulaSatisfied('=COUNTIF(INDIRECT("'+SHEET_LEVELS+'!A2:A"), E2)>0')
    .setAllowInvalid(false)
    .setHelpText('Pick a LevelID that exists in '+SHEET_LEVELS+'!A2:A')
    .build();
  sub.getRange('E2:E').setDataValidation(rule);

  // Widths (nice to have)
  [sub,lev,crit].forEach(function(sh){
    var cols = sh.getLastColumn();
    for (var c=1;c<=cols;c++){
      sh.autoResizeColumn(c);
    }
  });

  SpreadsheetApp.getUi().alert(
    'Sheets created.\n\n' +
    'Next:\n' +
    '1) Set GEMINI_API_KEY in Script properties (default provider).\n' +
    '2) (Optional) Set LLM_PROVIDER=openai and OPENAI_API_KEY.\n' +
    '3) Refresh, then use the Autograder menu.'
  );
}

function parseCriteriaTableCsv_() {
  // CSV is embedded into the script so setupSheets() can populate rubrics automatically.
  // Update it by replacing the string returned by getCriteriaTableCsvText_().
  var csvText = getCriteriaTableCsvText_();
  if (!csvText || !String(csvText).trim()) throw new Error('Embedded criteria table CSV is empty.');

  // Normalize newlines + trim.
  csvText = String(csvText).replace(/\r\n/g, '\n').replace(/\r/g, '\n').trim();

  // Apps Script's Utilities.parseCsv is surprisingly strict with some embedded-string CSV cases.
  // Use a tiny CSV parser here that supports:
  // - comma-separated fields
  // - quoted fields
  // - doubled quotes inside quoted fields ("" -> ")
  // - newlines inside quoted fields (not expected here, but safe)
  var values = parseCsvText_(csvText);
  if (!values || values.length < 2) throw new Error('criteria-table.csv has no data rows.');

  var header = values[0].map(function(h){ return String(h||'').trim(); });
  var idx = {};
  header.forEach(function(h,i){ idx[h]=i; });

  function col(name) { return idx[name]; }

  var byLevel = {};
  var outRows = [];
  for (var r=1; r<values.length; r++) {
    var row = values[r];
    if (!row || !row.length) continue;
    var levelId = String(row[col('LevelID')] || '').trim();
    if (!levelId) continue;
    byLevel[levelId] = true;
    outRows.push([
      levelId,
      String(row[col('CriterionID')] || '').trim(),
      row[col('Points')],
      String(row[col('Type')] || '').trim(),
      String(row[col('Description')] || '').trim(),
      String((col('Notes') !== undefined ? (row[col('Notes')]||'') : '') || '').trim(),
      String((col('Teacher Notes') !== undefined ? (row[col('Teacher Notes')]||'') : '') || '').trim()
    ]);
  }
  return { rows: outRows, byLevel: byLevel };
}

function parseCsvText_(text) {
  var rows = [];
  var row = [];
  var field = '';
  var inQuotes = false;

  for (var i = 0; i < text.length; i++) {
    var ch = text[i];

    if (inQuotes) {
      if (ch === '"') {
        // Doubled quote inside quoted field => literal quote.
        if (i + 1 < text.length && text[i + 1] === '"') {
          field += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        field += ch;
      }
      continue;
    }

    if (ch === '"') {
      inQuotes = true;
      continue;
    }
    if (ch === ',') {
      row.push(field);
      field = '';
      continue;
    }
    if (ch === '\n') {
      row.push(field);
      field = '';
      // Skip empty trailing line (common when text ends with newline)
      if (row.length > 1 || (row.length === 1 && String(row[0]).trim() !== '')) {
        rows.push(row);
      }
      row = [];
      continue;
    }
    field += ch;
  }

  // Flush last field/row.
  row.push(field);
  if (row.length > 1 || (row.length === 1 && String(row[0]).trim() !== '')) {
    rows.push(row);
  }

  // Guard against unclosed quote.
  if (inQuotes) throw new Error('CSV parse error: unmatched quote in embedded rubric (missing closing ").');

  return rows;
}

function getCriteriaTableCsvText_() {
  // IMPORTANT: This content mirrors the repo file `criteria-table.csv`.
  // If you update the CSV in the repo, update this string.
  return (
"LevelID,CriterionID,Points,Type,Description,Notes,Teacher Notes\n" +
"L3-08,ellipses_present,1,llm_check,\"The four orange ellipses are present in a 2x2 grid with 50 pixel spacing. Count it as OK if each ellipse is drawn while the current fill is orange (literal \"\"orange\"\" or a variable set to \"\"orange\"\").\",,\n" +
"L3-08,purple_rect_after_fill,4,llm_check,\"A rectangle is drawn after a call that sets fill to purple (either the literal \"\"purple\"\" or a variable assigned \"\"purple\"\"). The rectangle call may be rect(x,y) or rect(x,y,width,height); rectangle x and y coordinates must be within 10 pixels of the coordinates of the first ellipse (the ellipse in the top-left corner of the 2x2 grid)\",,\n" +
"L3-08,rect_after_ellipses,4,llm_check,The (purple) rectangle’s draw call appears later in the code than the four orange ellipse calls (so it is on top by draw order). It does not have to be the last shape in the program.,,\n" +
"L3-08,one_rect_present,1,llm_check,There is exactly one rectangle present (a single call of the rect() function),,\n" +
"L4-08,cloud_wider_than_tall,9,llm_check,\"After setting fill(\"\"white\"\") for the cloud, the cloud ellipse’s width is greater than its height (any numeric values OK as long as clearly wider).\",,\n" +
"L4-08,no_old_tall_cloud,1,llm_check,\"The original tall cloud ellipse at (150,100,100,200) is not left unchanged.\",,\n" +
"L5-07,left_uses_eyeSize,1,llm_check,The left eye draws an ellipse whose width and height are controlled by the variable eyeSize (not numeric literals).,,\n" +
"L5-07,right_uses_eyeSize,8,llm_check,The right eye draws an ellipse whose width and height are controlled by the variable eyeSize (not numeric literals).,,\n" +
"L5-07,eyeSize_declared,1,llm_check,The program declares a variable named eyeSize and assigns it a number.,,\n" +
"L5-07,pupil_bonus_scale,1,llm_check,The eyes contain pupils that consist of a second pair of ellipses is drawn in the same location as the eyes. The width and height of these pupil ellipses are always smaller than the eyes (either by dividing eyeSize by a number greater than 1 or multiplying eyeSize by a number less than 1). ,,\n" +
"L5-07,pupil_bonus_fill,1,llm_check,The pupil ellipses (the smaller ones drawn in the same location as the eyes) have a fill color different from the eyes.,,\n" +
"L5-06c,earSize_variable,2,llm_check,\"The ellipses that represent the bear's ears located at (130, 115) and (270, 115) have a width and height set equal to the earSize variable.\",,\n" +
"L5-06c,eyeSize_variable,2,llm_check,\"The ellipses that represent the bear's eyes located at (165, 175) and (235, 175) have a width and height set equal to the eyeSize variable.\",,\n" +
"L5-06c,x_pos_variable,2,llm_check,\"There is a variable that controls the x position of the entire bear. Any variable name is acceptable (center, xPos, etc). Look carefully and make sure that every x-coordinate of every ellipse, every line, and every arc is set correctly with reference to this variable. \" ,,\n" +
"L5-06c,y_pos_variable,2,llm_check,\"There is a variable that controls the y position of the entire bear. Any variable name is acceptable (yPos, yLoc, vert, etc). Look carefully and make sure that every y-coordinate of every ellipse, every line, and every arc is set correctly with reference to the variable. \" ,,\n" +
"L5-06c,scale_variable,2,llm_check,\"There is a multiplier variable that controls the scale of the entire bear. Any variable name is acceptable (bearSize, scale, size, etc.). Look carefully and make sure that every x-coordinate, y-coordinate, width, and height are properly multiplied by the variable to ensure proper placement and scale of all elements.\",,\n" +
"L6-07,at_least_3_new_circles,2,llm_check,The program draws at least six circles total (ellipse calls where width == height). This implies the student added at least three new circles beyond the original three.,,\n" +
"L6-07,all_unique_colors,2,llm_check,\"All circles (the original three plus the new three that the student added) have distinct fill colors. Accept named colors (e.g., \"\"green\"\") or color codes (e.g., \"\"#00ff00\"\", rgb(), hsl()).\",,\n" +
"L6-07,new_circles_random_y,3,llm_check,\"For every added circle, the ellipse’s Y coordinate is controlled by randomNumber(...).  Matching the numeric range of 190–210 is ideal, but credit should be awarded as long as the low value is not less than180 and the high value is not greater than 220.\",,\n" +
"L6-07,new_circles_correct_x,3,llm_check,\"Each added circle is exactly 40 pixels to the right of the previous one. The proper x-coordinates of all circles, from left to right, should be: 100, 140, 180, 220, 260, 300. If additional circles are optionally added, the pattern should continue (340, 380, 420, etc.)\",,\n" +
"L6-07adv,x_pos_variable,2,llm_check,\"There is a single variable that controls the horizontal origin of the entire caterpillar: changing it shifts all seven body circles (and their highlights) left/right together. Inter-circle spacing is derived from this origin and preserves the same layout at any overall size (i.e., spacing scales proportionally with size).\",,\n" +
"L6-07adv,y_pos_variable,2,llm_check,\"There is a variable that sets the vertical centerline for the caterpillar: changing it moves the whole row up/down, while each circle’s actual y stays within a small symmetric random band that scales proportionally with size (so the “wobble” looks the same at all scales).\",,\n" +
"L6-07adv,size_variable,2,llm_check,\"There is a global scale variable such that increasing it makes the caterpillar proportionally larger and decreasing it makes it smaller without changing the design’s proportions: circle diameters, inter-circle spacing, highlight size, highlight offset, and the random y-band all scale consistently with size.\",,\n" +
"L6-07adv,highlightSize_variable,2,llm_check,\"There is a variable that controls the size of the white highlight on each circle, and it scales proportionally with the overall size so the highlight effect looks the same at any caterpillar scale (implementation via ellipse/arc/stroke is fine).\",,\n" +
"L6-07adv,highlightOffset_variable,2,llm_check,\"There is a variable that controls how far each highlight is offset from its circle’s center, and this offset scales proportionally with size so the highlight sits in the same relative corner (e.g., up-left) at any scale; increasing/decreasing it moves the highlight farther/closer consistently across all circles.\",,\n" +
"L6-08a,x_pos_variable,2,llm_check,\"There is a variable that controls the x position of the entire set of concentric circles. Any variable name is acceptable (x, xPos, xPosition, etc). Look carefully and make sure that the x-coordinate of every ellipse is set correctly to this variable. \" ,,\n" +
"L6-08a,y_pos_variable,2,llm_check,\"There is a variable that controls the y position of the entire set of concentric circles. Any variable name is acceptable (y, yPos, yPosition, etc). Look carefully and make sure that the y-coordinate of every ellipse is set correctly to this variable. \" ,,\n" +
"L6-08a,size_variable,2,llm_check,\"There is a multiplier variable that controls the scale of the entire set of concentric circles. Any variable name is acceptable (scale, size, etc.). Look carefully and make sure that the width and height of every ellipse are properly multiplied by the variable to ensure proper scale.\",,\n" +
"L8-10,at_least_2_sprites,4,llm_check,At least two sprites are present in the program.,,\n" +
"L8-10,different_locations,2,llm_check,The sprites are in different locations (they don't have the exact same x and y coordinates).,,\n" +
"L8-10,sprite_animations,4,llm_check,All sprites have been assigned an animation using the sprite.setAnimation() method. The animation name (the string parameter used with the setAnimation method) should be different for each sprite.,,\n" +
"L9-05,burger_size,4,llm_check,burger sprite has been assigned a .scale property value of less than 1,,\n" +
"L9-05,fries_size,3,llm_check,fries sprite has been assigned a .scale property value of less than 1,,\n" +
"L9-05,dessert_size,3,llm_check,dessert sprite has been assigned a .scale property value of less than 1,,\n" +
"L9-05adv,x_pos_variable,2,llm_check,\"There is a variable that controls the x position of all elements. Any variable name is acceptable (x, xPos, xPosition, etc). The the ellipse (which represents the plate) should be located exactly at the xPos coordinate. The sprites (burger, fries, dessert) should be positioned relative to xPos using addition and subtraction. \" ,,\n" +
"L9-05adv,y_pos_variable,2,llm_check,\"There is a variable that controls the y position of all elements. Any variable name is acceptable (y, yPos, yPosition, etc). The the ellipse (which represents the plate) should be located exactly at the yPos coordinate. The sprites (burger, fries, dessert) should be positioned relative to yPos using addition and subtraction. \" ,,\n" +
"L9-05adv,scale_variable,2,llm_check,\"There is a multiplier variable that controls the scale of all elements, including the width and height of the ellipse, and the scale of the burger, fries, and dessert. This multiplier should also be used to scale the offset distance from xPos and yPos for each of the sprite locations (burger, fries, and dessert). Make sure that the multplier has been used everywhere necessary to ensure proper scaling and placement of all elements.\",,\n" +
"L10-05,two_texts,5,llm_check,\"The program draws at least two separate text elements using text(string, x, y), and each string is non-empty. Positions/colors are flexible.\",,\n" +
"L10-05,two_positions,5,llm_check,The two text elements appear at different positions (they don't have the same x / y coordinates).,,\n" +
"L10-05,text_stroke,2,llm_check,\"At least one text element is rendered in a high-contrast outlined style: either fill set to black with stroke set to white OR fill set to white with stroke set to black. Accept any equivalent color notation (e.g., \"\"black\"\"/\"\"white\"\", #000/#fff, #000000/#ffffff, rgb(0,0,0)/rgb(255,255,255), case-insensitive). The stroke/fill must be in effect when text(...) is drawn; strokeWeight ≥ 1 may be explicit or default.\",,\n" +
"L12-07,has_draw_loop,3,llm_check,The program defines a draw() function that runs continuously.,,\n" +
"L12-07,salt_y_randomized,3,llm_check,\"Inside draw(), the salt sprite’s Y position is randomized each frame (e.g., salt.y = randomNumber(...)). Any numeric range is fine.\",,\n" +
"L12-07,drawSprites_in_draw,2,llm_check,drawSprites() is called within draw() (not only once before the loop).,,\n" +
"L12-07,background_in_draw,2,llm_check,background(...) is called within draw() so the screen is cleared each frame (no trails).,,\n" +
"L12-07adv,shake_variable,5,llm_check,\"There is a variable that controls the shake distance of the salt sprite. Any variable name is fine (shake, strength, etc.). This variable should be used to affect the random number range that is used for the y property of the salt sprite. Addition, subtraction, or a combination of the two are acceptable (e.g.: randomNumber(200, 200 + shake) or randomNumber(200 - shake, 200), or randomNumber(200 - shake, 200 + shake)).\",,\n" +
"L13-07,all_fish_moving,4,llm_check,\"All three sprites (orangeFish, blueFish, greenFish) are moving (their x positions are changing over time within the draw() function)\",,\n" +
"L13-07,fish_different_speeds,4,llm_check,\"The three sprites (orangeFish, blueFish, greenFish) are all moving at different speeds (their x positions are changing by different values within the draw() function)\",,\n" +
"L13-07,blue_fish_faster,1,llm_check,The blueFish sprite is moving faster than the other two fish (its x position is changing more per frame than the other two),,\n" +
"L13-07,green_fish_slower,1,llm_check,The greenFish sprite is moving slower than the other two fish (its x position is changing less per frame than the other two),,\n" +
"L15-07,dinosaur_transforms,10,llm_check,\"There is an if-statement that checks to see whether the y position of the dinosaur sprite is less than a certain value (any number between 50 and 300 is acceptable), and if so, uses dinosaur.setAnimation(...) to change the dinosaur's animation\",,\n" +
"L16-06,flyer_movement_left,2.5,llm_check,\"There is an if-statement that checks to see whether the left arrow is being pressed (the letter 'a' is also acceptable) and if so, moves the sprite to the left by reducing its x position property. The if-statements can appear in any order (ignore all comments that indicate where each if-statement should appear)\",,\n" +
"L16-06,flyer_movement_right,2.5,llm_check,\"There is an if-statement that checks to see whether the right arrow is being pressed (the letter 'd' is also acceptable) and if so, moves the sprite to the right by increasing its x position property. The if-statements can appear in any order (ignore all comments that indicate where each if-statement should appear)\",,\n" +
"L16-06,flyer_movement_up,2.5,llm_check,\"There is an if-statement that checks to see whether the up arrow is being pressed (the letter 'w' is also acceptable) and if so, moves the sprite up by reducing its y position property. The if-statements can appear in any order (ignore all comments that indicate where each if-statement should appear)\",,\n" +
"L16-06,flyer_movement_down,2.5,llm_check,\"There is an if-statement that checks to see whether the down arrow is being pressed (the letter 's' is also acceptable) and if so, moves the sprite down by increasing its y position property. The if-statements can appear in any order (ignore all comments that indicate where each if-statement should appear)\",,\n" +
"L17-07,shake_when_clicking,4,llm_check,The creature only shakes while the mouse is being pressed (the line of code that randomizes the rotation property of the creature sprite should be inside of an if (mouseDown(...) conditional statement).,,\n" +
"L17-07,text_in_else_statement,4,llm_check,\"The text function that writes \"\"Press the mouse to shake the creature.\"\" to the screen runs only while the mouse is NOT being pressed (the text function should be placed in an else statement that runs when mouseDown(...) returns false)\",,\n" +
"L17-07,drawSprites_first,2,llm_check,The drawSprites() function must be called in the draw() loop BEFORE the if-else statement (otherwise the sprite background will cover the text and you won't be able to see it),,\n" +
"L19-09,fish_starts_with_key,2,llm_check,\"When the right arrow key is pressed, the fish starts moving to the right (there is an if-statement with keyWentDown(...) that sets fish.velocityX to a positive number. It's ok if the code responds to the left arrow key instead, or if both options are present.\",,\n" +
"L19-09,fish_moves_right_to_left,3,llm_check,\"When the fish reaches the right side of the screen, it starts moving left (there is an if-statement that checks whether fish.x is greater than 400, and if so, sets fish.velocityX to a negative number)\",,\n" +
"L19-09,fish_moves_left_to_right,3,llm_check,\"When the fish reaches the left side of the screen, it starts moving right (there is an if-statement that checks whether fish.x is less than 0, and if so, sets fish.velocityX to a positive number)\",,\n" +
"L19-09,fish_faces_move_dir,2,llm_check,\"The fish is always facing the direction that it's swimming (when fish.velocityX changes to a negative number, the fish's animation is set to \"\"fishL\"\", and when fish.velocityX changes to a positive number, the fish's animation is set to \"\"fishR\"\")\",,\n" +
"L20-07,horse_changes_to_unicorn,10,llm_check,\"When the rainbow sprite collides with the horse sprite, the horse's animation changes to a unicorn.\",,\n" +
"L21-GAME,background_in_draw,3,llm_check,\"Game has a background. This can be either a solid color using the background() function OR a sprite with a full-screen animation image assigned that is created before all others. If a sprite background is used, the background() function does NOT need to be present to receive credit.\",,\n" +
"L21-GAME,player_sprite,5,llm_check,\"Game has a player sprite. Ideally, the sprite is named 'player', but it may be called something else (e.g.: frog, alien, etc.)\",,\n" +
"L21-GAME,player_jump,4,llm_check,\"This is a jump button that cause the player to move upwards. Button can be up-arrow, spacebar, or w.\",,\n" +
"L21-GAME,player_ceiling,3,llm_check,The player sprite does not fly off of the top of the screen,,\n" +
"L21-GAME,player_floor,3,llm_check,The player sprite does not fall off of the bottom of the screen,,\n" +
"L21-GAME,obstacle_sprite,2,llm_check,\"Game has an obstacle sprite. Ideally, the sprite is named 'obstacle', but it may be called something else (e.g.: mushroom, spike, enemy, etc.)\",,\n" +
"L21-GAME,obstacle_movement,2,llm_check,\"The obstacle sprite has a negative velocityX value, causing it to constantly move from right to left\",,\n" +
"L21-GAME,obstacle_looping,2,llm_check,\"When the obstacle sprite reaches the left edge of the screen, it jumps back to the right edge of the screen, causing it to \"\"loop\"\" indefinitely\",,\n" +
"L21-GAME,obstacle_collision,2,llm_check,\"When the player touches the obstacle, the health variable goes down\",,\n" +
"L21-GAME,target_sprite,2,llm_check,\"Game has a collectible item sprite. Ideally, the sprite is named 'target', but it may be called something else (e.g.: fly, coin, etc.)\",,\n" +
"L21-GAME,target_sprite_movement,2,llm_check,\"The target sprite has a negative velocityX value, causing it to constantly move from right to left\",,\n" +
"L21-GAME,target_looping,2,llm_check,\"When the target sprite reaches the left edge of the screen, it jumps back to the right edge of the screen, causing it to \"\"loop\"\" indefinitely\",,\n" +
"L21-GAME,target_collision,2,llm_check,\"When the player touches the target, the score variable goes up\",,\n" +
"L21-GAME,score_display,2,llm_check,There is a score-counter displayed somewhere on the screen that properly displays the value of the 'score' variable,,\n" +
"L21-GAME,health_display,2,llm_check,There is a health-counter displayed somewhere on the screen that properly displays the value of the 'health' variable,,\n" +
"L21-GAME,game_over,2,llm_check,\"When the 'health' variable reaches 0, a Game Over message is displayed. This can be either a background() function with a text() function message, or a sprite-based image\",,\n" +
"L22-06,rocks_falls_back_down,10,llm_check,\"The rock, which is initially moving upward, slows its ascent and then falls back down. This is accomplished by increasing rock.velocityY each frame.\",,\n"
  );
}

/** === MENU === */
function onOpen() {
  var ui = SpreadsheetApp.getUi();

  var emailsMenu = ui.createMenu('Emails')
    .addItem('Email selected rows (if Email column set)', 'emailSelectedRows');

  var adminMenu = ui.createMenu('Admin')
    .addItem('Grade ALL rows (slow)', 'gradeAllRows')
    .addSeparator()
    .addItem('Import/backfill from Form Responses', 'backfillFormResponses')
    .addItem('Import/backfill + grade + email', 'backfillFormResponsesAndEmail');

  var diagnosticsMenu = ui.createMenu('Diagnostics')
  .addItem('Test LLM connection (current provider)', 'diagnosticsTestLLM')
  .addItem('Test structured JSON (current provider)', 'diagnosticsTestLLMStructured');

  ui.createMenu('Autograder')
    .addItem('Grade ungraded submissions (Score blank)', 'gradeNewRows')
    .addItem('Grade selected rows', 'gradeSelectedRows')
    .addSeparator()
    .addSubMenu(emailsMenu)
    .addSubMenu(adminMenu)
    .addSubMenu(diagnosticsMenu)
    .addSeparator()
    .addItem('Help / About', 'showAutograderHelp')
    .addToUi();
}

function showAutograderHelp() {
  var html = HtmlService.createHtmlOutput(
    '<div style="font-family: Arial, Helvetica, sans-serif; font-size: 13px; line-height: 1.4;">' +
    '<h2 style="margin:0 0 8px 0; font-size:16px;">Game Lab Autograder</h2>' +
    '<p style="margin:0 0 10px 0;">This spreadsheet grades Code.org Game Lab share links using rubric rows in the <b>Criteria</b> tab and level enable/model settings in <b>Levels</b>.</p>' +

    '<h3 style="margin:12px 0 6px 0; font-size:14px;">Common actions</h3>' +
    '<ul style="margin:0 0 10px 18px; padding:0;">' +
    '<li><b>Grade ungraded submissions</b>: grades rows where <code>Score</code> is blank.</li>' +
    '<li><b>Grade selected rows</b>: grades only the rows you highlight/select.</li>' +
    '</ul>' +

    '<h3 style="margin:12px 0 6px 0; font-size:14px;">Emails</h3>' +
    '<ul style="margin:0 0 10px 18px; padding:0;">' +
    '<li><b>Email selected rows</b>: sends a results email to the address in the <code>Email</code> column. Skips rows that already have <code>EmailedAt</code>.</li>' +
    '</ul>' +

    '<h3 style="margin:12px 0 6px 0; font-size:14px;">Admin</h3>' +
    '<ul style="margin:0 0 10px 18px; padding:0;">' +
    '<li><b>Grade ALL rows</b>: can be slow and may hit Apps Script quotas. Use only when needed.</li>' +
    '<li><b>Import/backfill from Form Responses</b>: copies a row from a Google Form response sheet into <b>Submissions</b>.</li>' +
    '<li><b>Import/backfill + grade + email</b>: backfills, grades, then emails. Use with care.</li>' +
    '</ul>' +

  '<h3 style="margin:12px 0 6px 0; font-size:14px;">Diagnostics</h3>' +
    '<ul style="margin:0 0 10px 18px; padding:0;">' +
  '<li><b>Test LLM connection</b>: verifies your currently selected provider and API key are set and reachable.</li>' +
  '<li><b>Test structured JSON</b>: verifies the grader can reliably parse strict JSON from the provider.</li>' +
    '</ul>' +

    '<h3 style="margin:12px 0 6px 0; font-size:14px;">Setup checklist</h3>' +
    '<ol style="margin:0 0 0 18px; padding:0;">' +
    '<li>Run <code>setupSheets()</code> once to create tabs + headers.</li>' +
  '<li>Set <code>GEMINI_API_KEY</code> (default) or <code>OPENAI_API_KEY</code> in Apps Script Project Settings → Script properties.</li>' +
  '<li>(Optional) Set <code>LLM_PROVIDER</code> to <code>gemini</code> or <code>openai</code>. Default is <code>gemini</code>.</li>' +
    '<li>Fill <b>Levels</b> + <b>Criteria</b> (see the repo’s <code>criteria-table.csv</code>).</li>' +
    '</ol>' +
    '</div>'
  ).setWidth(520).setHeight(560);

  SpreadsheetApp.getUi().showModalDialog(html, 'Autograder Help');
}

function diagnosticsTestLLM() {
  var p = getLLMProvider_();
  if (p === 'openai') return pingGPT();
  return pingGemini_();
}

function diagnosticsTestLLMStructured() {
  var p = getLLMProvider_();
  if (p === 'openai') return pingGPTStructured();
  return pingGeminiStructured_();
}



/** === PUBLIC ACTIONS === */
function gradeNewRows() {
  var sh = getSheet_(SHEET_SUB);
  var data = sh.getDataRange().getValues();
  var head = headers_(data[0]);
  var targets = [];
  for (var r=1; r<data.length; r++) {
    var score = data[r][head.Score], url = data[r][head.ShareURL], lvl = data[r][head.LevelID];
    if (url && lvl && (score === '' || score === null)) targets.push(r+1);
  }
  gradeRows_(targets);
}
function gradeSelectedRows() {
  var sh = getSheet_(SHEET_SUB);
  var sel = sh.getActiveRange();
  var rows = [];
  for (var r=sel.getRow(); r<sel.getRow()+sel.getNumRows(); r++) rows.push(r);
  gradeRows_(rows);
}
function gradeAllRows() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.alert(
    'Grade ALL rows?',
    'This can be slow and may hit Apps Script quotas if you have many submissions.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  if (res !== ui.Button.YES) return;

  var sh = getSheet_(SHEET_SUB);
  var last = sh.getLastRow(), rows = [];
  for (var r=2; r<=last; r++) rows.push(r);
  gradeRows_(rows);
}

/** === CORE GRADING === */
function gradeRows_(rowNums) {
  if (!rowNums || !rowNums.length) return;

  var sh = getSheet_(SHEET_SUB);
  var data = sh.getDataRange().getValues();
  var head = headers_(data[0]);
  var critByLevel = loadCriteriaByLevel_();
  var enabledLevels = loadEnabledLevels_();

  rowNums.forEach(function(rowNum) {
    try {
      var r = rowNum - 1;
      var url = (data[r][head.ShareURL] || '').toString().trim();
      var levelId = (data[r][head.LevelID] || '').toString().trim();

      if (!url || !levelId) {
        writeRow_(sh, rowNum, head, { Status: 'No URL/LevelID', Score: 0, Notes: '' });
        return;
      }
      if (!enabledLevels[levelId]) {
        writeRow_(sh, rowNum, head, { Status: 'Level disabled/unknown', Score: 0, Notes: '' });
        return;
      }

      var crits = critByLevel[levelId] || [];
      var maxPts = crits.reduce(function(s,c){ return s + (Number(c.Points)||0); }, 0);
      if (!crits.length) {
        writeRow_(sh, rowNum, head, { Status: 'No criteria found', Score: 0, MaxScore: 0, Notes: '' });
        return;
      }

      var channelId = extractChannelId_(url);
      if (!channelId) {
        writeRow_(sh, rowNum, head, {
          ChannelID: '',
          Score: 0,
          MaxScore: maxPts,
          Status: 'Invalid share link (no ChannelID)',
          Notes: 'Expecting a projects/gamelab/<id> share URL'
        });
        return;
      }

      var fetched = fetchGameLabSource_(channelId);
      if (!fetched || !fetched.ok) {
        writeRow_(sh, rowNum, head, {
          ChannelID: channelId,
          Score: 0,
          MaxScore: maxPts,
          Status: 'Invalid share link or unreadable project',
          Notes: (fetched && fetched.msg) ? fetched.msg : 'Fetch failed'
        });
        return;
      }

      // Grade (may hit cache)
      var res = runCriteria_(fetched.src, crits, levelId);
      var patch = {
        ChannelID: channelId,
        Score: res.score,
        MaxScore: res.max,
        Status: 'OK',
        Notes: res.notes.join(' | ')
      };

      // Compare against existing row; if unchanged AND (hit cache), skip writing
      var prev = {
        ChannelID: data[r][head.ChannelID],
        Score:     data[r][head.Score],
        MaxScore:  data[r][head.MaxScore],
        Status:    data[r][head.Status],
        Notes:     data[r][head.Notes]
      };
      var changed = Object.keys(patch).some(function(k){
        return String(prev[k]||'') !== String(patch[k]||'');
      });

      // We can detect cache hit by peeking into runCriteria_ → openaiGrade_ return.
      // Add a tiny signal: if ANY note starts with ✅/❌ (normal), we can't tell.
      // So: we expose a global lastCacheHit flag set by openaiGrade_ (optional),
      // or simpler: treat "no change" as "no write".
      if (!changed) {
        // nothing changed — avoid rewriting cells
        return;
      }

      writeRow_(sh, rowNum, head, patch);

    } catch (e) {
      writeRow_(sh, rowNum, head, { Status: 'Error', Notes: String(e) });
    }
  });
}


/** === LLM ENGINE (Responses API with robust fallback) === */
function runCriteria_(src, crits, levelIdOpt) {
  // Split local vs LLM checks
  var localCrits = [], llmCrits = [];
  crits.forEach(function(c){
    var t = String(c.Type||'').trim().toLowerCase();
    if (t === 'code_nonempty' || t === 'contains' || t === 'regex_present' || t === 'regex_absent') localCrits.push(c);
    else if (t === 'llm_check') llmCrits.push(c);
  });

  var total = crits.reduce(function(s,c){ return s + (Number(c.Points)||0); }, 0);
  var got = 0, notes = [];

  // ---- Local checks (no LLM; deterministic)
  localCrits.forEach(function(c, i){
    var id = String(c.CriterionID || ('L'+i));
    var pts = Number(c.Points)||0;
    var t   = String(c.Type||'').trim().toLowerCase();
    var desc= String(c.Description||'');
    var pass = false, reason = '';

    if (t === 'code_nonempty') {
      pass = !!(src && String(src).trim().length >= 10); // tweak threshold if you want
      reason = pass ? '' : 'Code appears empty/too short.';
    } else if (t === 'contains') {
      // Case-insensitive substring match on Description
      var needle = desc.trim();
      pass = !!(needle && String(src||'').toLowerCase().indexOf(needle.toLowerCase()) >= 0);
      reason = pass ? '' : ('Missing text: ' + needle);
    } else if (t === 'regex_present') {
      try {
        var re = new RegExp(desc, 'i');
        pass = re.test(String(src||''));
        reason = pass ? '' : 'Pattern not found.';
      } catch (e) {
        pass = false; reason = 'Bad regex.';
      }
    } else if (t === 'regex_absent') {
      try {
        var re2 = new RegExp(desc, 'i');
        pass = !re2.test(String(src||'')); // pass if NOT present
        reason = pass ? '' : 'Forbidden pattern present.';
      } catch (e2) {
        pass = false; reason = 'Bad regex.';
      }
    }

    if (pass) got += pts;
    notes.push((pass ? '✅ ' : '❌ ') + (c.Description || id) + (pass ? '' : (reason ? ' – ' + reason : '')));
  });

  // ---- LLM checks (only if any)
  if (llmCrits.length) {
    var levelId = levelIdOpt || (llmCrits[0] && llmCrits[0].LevelID) || '';
  var res = llmGrade_(levelId, src, llmCrits);
    var byId = res.byId || {};
    llmCrits.forEach(function(c, i){
      var id = String(c.CriterionID || ('C'+i));
      var pts = Number(c.Points)||0;
      var r = byId[id] || { pass:false, reason:'' };
      if (r.pass) got += pts;
      notes.push((r.pass ? '✅ ' : '❌ ') + (c.Description || id) + (r.pass ? '' : (r.reason ? ' – ' + r.reason : '')));
    });
  }

  return { score: got, max: total, notes: notes };
}

function buildRubricPrompt_(levelId, src, llmCrits) {
  var checks = llmCrits.map(function(c, i){
    return {
      id: String(c.CriterionID || ('C'+i)),
      description: String(c.Description || '').trim(),
      points: Number(c.Points)||0
    };
  });

  var system =
    'You are a strict, consistent autograder for Code.org Game Lab (p5.js-style JavaScript). ' +
    'Given student code and rubric checks, decide PASS/FAIL for each check. ' +
    'If the code is empty/unreadable, mark all FAIL and set unreadable=true. ' +
    'Output JSON only.';

  var scoringRules = [
    'Treat code order as draw order: later shapes appear on top.',
    'rect(x,y) or rect(x,y,w,h) are both valid.',
    'Color can be literal ("purple") or a variable assigned that literal.',
    'Whitespace, comments, and semicolons are irrelevant.',
    'If unsure, mark FAIL (false).'
  ].join('\n- ');

  var user =
    'LEVEL: ' + levelId + '\n' +
    'SCORING RULES:\n- ' + scoringRules + '\n\n' +
    'Return ONLY JSON with this shape: {"unreadable":boolean,"checks":[{"id":string,"pass":boolean,"reason":string}]}.\n' +
    'CHECKS (IDs and descriptions):\n' +
    checks.map(function(x){ return '- ' + x.id + ': ' + x.description + ' (points ' + x.points + ')'; }).join('\n') +
    '\n\nCODE (fenced):\n```javascript\n' + (src||'') + '\n```';

  return { system: system, user: user, expectedIds: checks.map(function(x){ return x.id; }) };
}

function llmGrade_(levelId, src, llmCrits) {
  var provider = getLLMProvider_();
  if (provider === 'openai') return openaiGrade_(levelId, src, llmCrits);
  return geminiGrade_(levelId, src, llmCrits);
}

function openaiGrade_(levelId, src, llmCrits) {
  var key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) throw new Error('Missing OPENAI_API_KEY in Script properties');

  var built = buildRubricPrompt_(levelId, src, llmCrits);
  var expectedIds = built.expectedIds;

  var schema = {
    name: 'autograde_result',
    schema: {
      type: 'object',
      additionalProperties: false,
      properties: {
        unreadable: { type: 'boolean' },
        checks: {
          type: 'array',
          items: {
            type: 'object',
            additionalProperties: false,
            required: ['id','pass'],
            properties: { id:{type:'string'}, pass:{type:'boolean'}, reason:{type:'string'} }
          }
        }
      },
      required: ['checks']
    },
    strict: true
  };

  var model = getModelForLevel_(levelId) || getDefaultModel_();
  var result = callResponsesStructured_(model, key, built.system, built.user, schema);
  var parsed = normalizeAutogradeJson_(result.text, expectedIds);

  var byId = {};
  (parsed.checks || []).forEach(function(ch){
    byId[String(ch.id)] = { pass: !!ch.pass, reason: ch.reason || '' };
  });
  return { byId: byId, raw: parsed, provider: 'openai', model: model };
}

function geminiGrade_(levelId, src, llmCrits) {
  var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) throw new Error('Missing GEMINI_API_KEY in Script properties');

  var built = buildRubricPrompt_(levelId, src, llmCrits);
  var expectedIds = built.expectedIds;
  var model = getModelForLevel_(levelId) || getDefaultModel_();

  // Gemini API: generateContent
  // Docs vary by model/version; this uses the widely-supported v1beta endpoint.
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(key);

  // We put system+user into a single user message to keep this simple and portable.
  var body = {
    contents: [
      {
        role: 'user',
        parts: [{ text: built.system + '\n\n' + built.user }]
      }
    ],
    generationConfig: {
      temperature: 0,
      topP: 1
    }
  };

  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    muteHttpExceptions: true,
    payload: JSON.stringify(body)
  });

  var code = resp.getResponseCode();
  var txt = resp.getContentText();
  if (code >= 400) throw new Error('Gemini HTTP ' + code + ': ' + txt);

  var outText = extractGeminiText_(txt);
  var parsed = normalizeAutogradeJson_(outText, expectedIds);

  var byId = {};
  (parsed.checks || []).forEach(function(ch){
    byId[String(ch.id)] = { pass: !!ch.pass, reason: ch.reason || '' };
  });

  return { byId: byId, raw: parsed, provider: 'gemini', model: model };
}

function extractGeminiText_(txt) {
  var obj = {}; try { obj = JSON.parse(txt); } catch(e) { return ''; }
  // Typical shape: candidates[0].content.parts[0].text
  var c = obj && obj.candidates && obj.candidates[0];
  var parts = c && c.content && c.content.parts;
  if (parts && parts.length) {
    return parts.map(function(p){ return p.text || ''; }).join('');
  }
  // Fallback: some responses use "text" directly
  return (obj && obj.text) ? String(obj.text) : '';
}



/** === Robust Responses helpers === */
function extractResponsesText_(txt) {
  var obj = {}; try { obj = JSON.parse(txt); } catch(e) { return ''; }
  return (
    (obj && obj.output_text) ||
    (obj && obj.output && obj.output[0] && obj.output[0].content && obj.output[0].content[0] && obj.output[0].content[0].text) ||
    (obj && obj.choices && obj.choices[0] && obj.choices[0].message && obj.choices[0].message.content) ||
    ''
  );
}
// Attempts: json_schema → json_object → plain JSON

// Attempts: json_schema → json_object → plain JSON, all at temperature 0
function callResponsesStructured_(model, key, system, user, schema) {
  function fetchBody(body) {
    return UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
      method: 'post',
      contentType: 'application/json',
      muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + key },
      payload: JSON.stringify(body)
    });
  }

  // Shared base for all attempts (deterministic)
  var base = {
    model: model,
    input: [{ role: 'system', content: system }, { role: 'user', content: user }],
    temperature: 0,
    top_p: 1
  };

  // 1) JSON Schema
  var b1 = JSON.parse(JSON.stringify(base));
  b1.response_format = { type: 'json_schema', json_schema: schema };
  var resp1 = fetchBody(b1), code1 = resp1.getResponseCode(), txt1 = resp1.getContentText();
  if (code1 < 400) {
    var t1 = extractResponsesText_(txt1);
    try { JSON.parse(t1); return { code: code1, text: t1, usedModel: model }; } catch(_) {}
  }

  // 2) json_object
  var b2 = JSON.parse(JSON.stringify(base));
  b2.response_format = { type: 'json_object' };
  b2.input[1].content = user + '\n\nReturn a JSON object with this exact shape: ' +
    '{"unreadable":boolean,"checks":[{"id":string,"pass":boolean,"reason":string}]}';
  var resp2 = fetchBody(b2), code2 = resp2.getResponseCode(), txt2 = resp2.getContentText();
  if (code2 < 400) {
    var t2 = extractResponsesText_(txt2);
    try { JSON.parse(t2); return { code: code2, text: t2, usedModel: model }; } catch(_) {}
  }

  // 3) Plain (no response_format) but “ONLY JSON”
  var b3 = JSON.parse(JSON.stringify(base));
  b3.input[1].content = user + '\n\nReturn ONLY JSON, no prose.';
  var resp3 = fetchBody(b3), code3 = resp3.getResponseCode(), txt3 = resp3.getContentText();
  var t3 = extractResponsesText_(txt3);
  return { code: code3, text: t3, usedModel: model };
}



/** === LEVELS & SHEET HELPERS === */
function loadCriteriaByLevel_() {
  var sh = getSheet_(SHEET_CRIT);
  var values = sh.getDataRange().getValues();
  var head = headers_(values[0]);
  var map = {};
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var levelId = (row[head.LevelID] || '').toString().trim();
    if (!levelId) continue;
    (map[levelId] = map[levelId] || []).push({
      LevelID: levelId,
      CriterionID: row[head.CriterionID],
      Points: row[head.Points],
      Type: row[head.Type],
      Description: row[head.Description],
  Notes: (head.Notes !== undefined ? row[head.Notes] : ''),
  TeacherNotes: (head['Teacher Notes'] !== undefined ? row[head['Teacher Notes']] : '')
    });
  }
  return map;
}
function loadEnabledLevels_(){
  var sh=getSheet_(SHEET_LEVELS), vals=sh.getDataRange().getValues(), head=headers_(vals[0]), set={};
  for (var i=1;i<vals.length;i++){
    var id=(vals[i][head.LevelID]||'').toString().trim();
    var en=String(vals[i][head.Enabled]).toUpperCase()!=='FALSE';
    if(id) set[id]=en;
  }
  return set;
}
function getModelForLevel_(levelId) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LEVELS);
  if (!sh) return '';
  var vals = sh.getDataRange().getValues();
  var head = headers_(vals[0]);
  for (var i=1;i<vals.length;i++){
    var id=(vals[i][head.LevelID]||'').toString().trim();
    if (id===levelId) {
      var model=(head.Model!==undefined ? (vals[i][head.Model]||'') : '').toString().trim();
    return model || '';
    }
  }
  return '';
}
function getSheet_(name){ var sh=SpreadsheetApp.getActive().getSheetByName(name); if(!sh) throw new Error('Missing sheet: '+name); return sh; }
function headers_(row1){ var m={}; for (var i=0;i<row1.length;i++) m[row1[i]]=i; return m; }
function writeRow_(sh,rowNum,head,patch){
  var row=sh.getRange(rowNum,1,1,sh.getLastColumn()).getValues()[0];
  Object.keys(patch).forEach(function(k){ if(head[k]!==undefined) row[head[k]]=patch[k]; });
  sh.getRange(rowNum,1,1,row.length).setValues([row]);
}

/** === CODE.ORG FETCH HELPERS === */
function extractChannelId_(url){
  var m = String(url).match(/https?:\/\/studio\.code\.org\/projects\/gamelab\/([A-Za-z0-9\-_]+)/i);
  return m ? m[1] : '';
}
function fetchGameLabSource_(channelId){
  var u='https://studio.code.org/v3/sources/'+encodeURIComponent(channelId)+'/main.json';
  var res=UrlFetchApp.fetch(u,{muteHttpExceptions:true});
  var code=res.getResponseCode(), body=res.getContentText();
  if(code>=400) return {ok:false, src:'', msg:'HTTP '+code+' from Code.org'};
  try{
    var parsed = JSON.parse(body);
    var src = (typeof parsed==='string') ? parsed :
              (parsed && (parsed.source || parsed.code)) ? (parsed.source || parsed.code) : '';
    if(!src || src.trim().length<10) return {ok:false, src:'', msg:'Empty or too-short source'};
    return {ok:true, src:src, msg:'OK'};
  }catch(e){
    return {ok:false, src:'', msg:'Non-JSON response (likely invalid share link)'};
  }
}

/** === EMAIL WORKFLOW (optional) === */
function onFormSubmit(e) {
  try {
    var subSh = getSheet_(SHEET_SUB); // "Submissions"
    var subHeaders = subSh.getRange(1, 1, 1, subSh.getLastColumn()).getValues()[0];
    var subMap = headers_(subHeaders);

    // Source = the tab the Form writes to (e.g., "Form Responses 1")
    var srcSh = e.range.getSheet();
    var srcHeaders = srcSh.getRange(1, 1, 1, srcSh.getLastColumn()).getValues()[0];
    var srcMap = headersSmart_(srcHeaders);

    // Grab the submitted row’s values
    var values = (e.values && e.values.length === srcHeaders.length)
      ? e.values
      : srcSh.getRange(e.range.getRow(), 1, 1, srcSh.getLastColumn()).getValues()[0];

    function getField(name, fallback) {
      var idx = srcMap[name];
      return (idx !== undefined) ? values[idx] : (fallback || '');
    }

    var out = new Array(subHeaders.length).fill('');
    if (subMap.Timestamp !== undefined) out[subMap.Timestamp] = getField('Timestamp', new Date());
    if (subMap.Email     !== undefined) out[subMap.Email]     = getField('Email', '');
    if (subMap.First     !== undefined) out[subMap.First]     = getField('First', '');
    if (subMap.Last      !== undefined) out[subMap.Last]      = getField('Last', '');
    if (subMap.Period    !== undefined) out[subMap.Period]    = getField('Period', '');
    if (subMap.LevelID   !== undefined) out[subMap.LevelID]   = getField('LevelID', '');
    if (subMap.ShareURL  !== undefined) out[subMap.ShareURL]  = getField('ShareURL', '');

    // Append normalized row into Submissions
    subSh.appendRow(out);

    // Grade the new row and (optionally) email
    var newRow = subSh.getLastRow();
    gradeRows_([newRow]);
    sendEmailForRow_(newRow); // safe to keep; it no-ops if Email column empty

  } catch (err) {
    Logger.log('onFormSubmit error: ' + err);
  }
}



function emailSelectedRows() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SUB);
  var sel = sh.getActiveRange();
  for (var r = sel.getRow(); r < sel.getRow() + sel.getNumRows(); r++) sendEmailForRow_(r);
}
function sendEmailForRow_(rowNum) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SUB);
  var data = sh.getDataRange().getValues();
  var head = headers_(data[0]);

  if (head.Email === undefined) return;
  if (head.EmailedAt === undefined) head.EmailedAt = -1;

  var row = sh.getRange(rowNum, 1, 1, sh.getLastColumn()).getValues()[0];
  var email = (row[head.Email] || '').toString().trim();
  if (!email) return;
  if (head.EmailedAt >= 0 && row[head.EmailedAt]) return;

  var first = row[head.First] || '';
  var last  = row[head.Last] || '';
  var level = row[head.LevelID] || '';
  var url   = row[head.ShareURL] || '';
  var score = row[head.Score] || 0;
  var max   = row[head.MaxScore] || 0;
  var status= row[head.Status] || '';
  var notes = (row[head.Notes] || '').toString();
  var who = [first,last].filter(Boolean).join(' ').trim() || 'Student';

  var subject = '[Autograder] ' + level + ' — ' + score + '/' + max + (who ? (' — ' + who) : '');
  var items = notes ? notes.split(' | ') : [];
  var htmlNotes = items.length ? ('<ul>' + items.map(function(x){ return '<li>' + esc_(x) + '</li>'; }).join('') + '</ul>') : '<em>No detailed notes.</em>';
  var statusMsg = (status === 'OK') ? 'Your submission was graded automatically.' : 'Your submission could not be fully graded yet: <strong>' + esc_(status) + '</strong>.';

  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif;font-size:14px;">' +
    '<p>Hi ' + esc_(who) + ',</p>' +
    '<p>' + statusMsg + '</p>' +
    '<p><strong>Level:</strong> ' + esc_(level) + '<br>' +
    '<strong>Score:</strong> ' + esc_(score + '/' + max) + '<br>' +
    (url ? ('<strong>Link:</strong> <a href="' + esc_(url) + '">your project</a>') : '') +
    '</p>' +
    '<p><strong>Checks:</strong></p>' + htmlNotes +
    '<p style="color:#666;">This email was generated by the class autograder.</p>' +
    '</div>';

  var text =
    'Hi ' + who + ',\n\n' +
    ((status === 'OK') ? 'Your submission was graded automatically.\n' : 'Your submission could not be fully graded: ' + status + '\n') +
    '\nLevel: ' + level +
    '\nScore: ' + score + '/' + max +
    (url ? '\nLink: ' + url : '') +
    '\n\nChecks:\n- ' + (items.length ? items.join('\n- ') : 'No detailed notes.') +
    '\n\n(This email was generated by the class autograder.)';

  GmailApp.sendEmail(email, subject, text, {htmlBody: html});
  if (head.EmailedAt >= 0) sh.getRange(rowNum, head.EmailedAt + 1).setValue(new Date());
}

/** === PING MENU FUNCTIONS === */
function pingGPT() {
  try {
    var key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!key) { SpreadsheetApp.getUi().alert('OPENAI_API_KEY is not set in Script properties.'); return; }
    var model = (typeof DEFAULT_MODEL === 'string' && DEFAULT_MODEL) ? DEFAULT_MODEL : 'gpt-5-mini';
    var body = { model: model, input: [{ role: 'user', content: 'Reply with the single word: pong.' }] };
    var resp = UrlFetchApp.fetch('https://api.openai.com/v1/responses', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + key }, payload: JSON.stringify(body)
    });
    var code = resp.getResponseCode(), txt  = resp.getContentText();
    var out = extractResponsesText_(txt);
    SpreadsheetApp.getUi().alert('Ping GPT → HTTP ' + code + ' on ' + model + (out ? '\n\nOutput: ' + out : '\n\nSee Logs for details.'));
  } catch (err) {
    SpreadsheetApp.getUi().alert('Ping failed: ' + err);
  }
}

function pingGemini_() {
  try {
    var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) { SpreadsheetApp.getUi().alert('GEMINI_API_KEY is not set in Script properties.'); return; }
    var model = getDefaultModel_();
    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(key);
    var body = {
      contents: [{ role: 'user', parts: [{ text: 'Reply with the single word: pong.' }] }],
      generationConfig: { temperature: 0, topP: 1 }
    };
    var resp = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify(body)
    });
    var code = resp.getResponseCode();
    var txt = resp.getContentText();
    var out = extractGeminiText_(txt);
    SpreadsheetApp.getUi().alert('Ping Gemini → HTTP ' + code + ' on ' + model + (out ? '\n\nOutput: ' + out : '\n\nSee Logs for details.'));
  } catch (err) {
    SpreadsheetApp.getUi().alert('Gemini ping failed: ' + err);
  }
}

function pingGeminiStructured_() {
  try {
    var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) { SpreadsheetApp.getUi().alert('GEMINI_API_KEY is not set in Script properties.'); return; }
    var model = getDefaultModel_();

    var checks = [
      { id: 'has_purple', description: 'Code sets fill("purple") before drawing a rectangle.' },
      { id: 'has_draw',   description: 'Code defines a draw() function.' }
    ];

    var system = 'You are a strict autograder. Decide PASS/FAIL per check. Output JSON only.';
    var user =
      'Return ONLY JSON with this exact shape: {"unreadable":boolean,"checks":[{"id":string,"pass":boolean,"reason":string}]}\n\n' +
      'CHECKS:\n' + checks.map(function(c){ return '- ' + c.id + ': ' + c.description; }).join('\n') +
      '\n\nCODE:\n```javascript\nfill("purple"); rect(10,10,20,20);\n```';

    var url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent?key=' + encodeURIComponent(key);
    var body = {
      contents: [{ role: 'user', parts: [{ text: system + '\n\n' + user }] }],
      generationConfig: { temperature: 0, topP: 1 }
    };

    var resp = UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify(body)
    });
    var code = resp.getResponseCode();
    var txt = resp.getContentText();
    if (code >= 400) {
      SpreadsheetApp.getUi().alert('Structured ping → HTTP ' + code + ' on ' + model + '\n\nSee Logs for response body.');
      Logger.log('Gemini error body:\n%s', txt);
      return;
    }

    var text = extractGeminiText_(txt);
    var parsed = normalizeAutogradeJson_(text, checks.map(function(c){ return c.id; }));
    var ok = parsed && Array.isArray(parsed.checks) && parsed.checks.length;

    SpreadsheetApp.getUi().alert(
      'Structured ping → HTTP ' + code + ' on ' + model +
      (ok ? '\n\nParsed checks: ' + parsed.checks.length : '\n\nCould not parse structured JSON (see Logs).')
    );
    Logger.log('Raw text:\n%s', text);
    Logger.log('Parsed object:\n%s', JSON.stringify(parsed));
  } catch (err) {
    SpreadsheetApp.getUi().alert('Gemini structured ping failed: ' + err);
  }
}

function pingGPTStructured() {
  try {
    var key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    if (!key) { SpreadsheetApp.getUi().alert('OPENAI_API_KEY is not set in Script properties.'); return; }

    var model = (typeof DEFAULT_MODEL === 'string' && DEFAULT_MODEL) ? DEFAULT_MODEL : 'gpt-5-mini';

    var checks = [
      { id: 'has_purple', description: 'Code sets fill("purple") before drawing a rectangle.' },
      { id: 'has_draw',   description: 'Code defines a draw() function.' }
    ];

    var system = 'You are a strict autograder. Decide PASS/FAIL per check. Output JSON only per schema.';
    var user =
      'CHECKS:\n' + checks.map(function(c){ return '- ' + c.id + ': ' + c.description; }).join('\n') +
      '\n\nCODE:\n```javascript\nfill("purple"); rect(10,10,20,20);\n```';

    var schema = {
      name: 'autograde_result',
      schema: {
        type: 'object',
        additionalProperties: false,
        properties: {
          unreadable: { type: 'boolean' },
          checks: {
            type: 'array',
            items: {
              type: 'object',
              additionalProperties: false,
              required: ['id','pass'],
              properties: { id:{type:'string'}, pass:{type:'boolean'}, reason:{type:'string'} }
            }
          }
        },
        required: ['checks']
      },
      strict: true
    };

    // Robust call: json_schema → json_object → plain JSON
    var result = callResponsesStructured_(model, key, system, user, schema);
    var code   = result.code;
    var used   = result.usedModel || model;
    var text   = result.text;

    // Normalize whatever came back into {unreadable, checks:[{id,pass,reason}]}
    var parsed = normalizeAutogradeJson_(text, checks.map(function(c){ return c.id; }));
    var ok = parsed && Array.isArray(parsed.checks) && parsed.checks.length;

    SpreadsheetApp.getUi().alert(
      'Structured ping → HTTP ' + code + ' on ' + used +
      (ok ? '\n\nParsed checks: ' + parsed.checks.length : '\n\nCould not parse structured JSON (see Logs).')
    );
    Logger.log('Raw text:\n%s', text);
    Logger.log('Parsed object:\n%s', JSON.stringify(parsed));
  } catch (err) {
    SpreadsheetApp.getUi().alert('Structured ping failed: ' + err);
  }
}

/** === UTIL === */
function esc_(s){
  return String(s).replace(/[&<>\"']/g, function(c){
    return {"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[c];
  });
}

// Strip ```json … ``` fences if the model wraps its JSON in a code block
function stripCodeFences_(s) {
  s = String(s || '').trim();
  if (s.startsWith('```')) {
    s = s.replace(/^```[\w-]*\s*/i, '').replace(/\s*```$/,'');
  }
  return s.trim();
}

// Normalize PASS/FAIL/true/false/yes/no to boolean
function toBool_(v) {
  if (typeof v === 'boolean') return v;
  var s = String(v || '').trim().toLowerCase();
  return s === 'pass' || s === 'passed' || s === 'true' || s === 'yes' || s === 'y' || s === '1';
}

/**
 * Accepts multiple response shapes and returns:
 *   { unreadable:boolean, checks:[{id, pass, reason}] }
 * - Supports:
 *   1) {checks:[{id, pass, reason}], unreadable?}
 *   2) {results:{id:{pass,reason}, ...}}  or  {id:"PASS"/"FAIL", ...}
 *   3) Plain object keyed by check IDs (like your log)
 */
function normalizeAutogradeJson_(text, expectedIds) {
  var out = { unreadable: false, checks: [] };
  text = stripCodeFences_(text);
  var obj;
  try { obj = JSON.parse(text); } catch (e) { return out; }

  // Preferred schema
  if (obj && Array.isArray(obj.checks)) {
    out.unreadable = !!obj.unreadable;
    obj.checks.forEach(function(ch){
      if (!ch) return;
      out.checks.push({
        id: String(ch.id),
        pass: toBool_(ch.pass),
        reason: ch.reason ? String(ch.reason) : ''
      });
    });
    return out;
  }

  // Maybe nested under "results"
  var source = (obj && typeof obj.results === 'object' && obj.results) ? obj.results : obj;

  // If the response is a map keyed by ids, convert it
  var ids = expectedIds && expectedIds.length ? expectedIds : Object.keys(source || {});
  ids.forEach(function(id){
    if (!source || !(id in source)) return;
    var v = source[id], pass = false, reason = '';
    if (v && typeof v === 'object') {
      if ('pass' in v) pass = toBool_(v.pass);
      else if ('result' in v) pass = toBool_(v.result);
      else if ('ok' in v) pass = toBool_(v.ok);
      if (v.reason) reason = String(v.reason);
    } else {
      pass = toBool_(v);
    }
    out.checks.push({ id: String(id), pass: pass, reason: reason });
  });

  return out;
}

function sha256_(s) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s);
  return raw.map(function(b){ var v=(b<0?b+256:b); return ('0'+v.toString(16)).slice(-2); }).join('');
}
function getGradeCache_(key){ return PropertiesService.getDocumentProperties().getProperty(key); }
function setGradeCache_(key, value){ PropertiesService.getDocumentProperties().setProperty(key, value); }
function clearGradeCache(){
  var props=PropertiesService.getDocumentProperties(), all=props.getProperties(), n=0;
  Object.keys(all).forEach(function(k){ if (k.indexOf('grade:')===0){ props.deleteProperty(k); n++; }});
  SpreadsheetApp.getUi().alert('Cleared '+n+' cached grade entries.');
}

function inspectL507() {
  inspectLevel('L5-07');
}

function inspectLevel(levelId) {
  var critByLevel = loadCriteriaByLevel_();
  var crits = critByLevel[levelId] || [];
  var model = getModelForLevel_(levelId) || (typeof DEFAULT_MODEL==='string' ? DEFAULT_MODEL : '(none)');

  var summary = [
    'Level: ' + levelId,
    'Model used: ' + model,
    'Criteria count: ' + crits.length
  ];

  var lines = crits.map(function(c, i){
    return (i+1) + ') id=' + c.CriterionID +
           ' | points=' + c.Points +
           ' | type=' + c.Type +
           ' | desc="' + String(c.Description).slice(0,80).replace(/\n/g,' ') + (String(c.Description).length>80?'…':'') + '"';
  });

  Logger.log(summary.join('\n'));
  Logger.log('--- Criteria ---\n' + (lines.join('\n') || '(none)'));
}

function debugSelectedRow() {
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SUB);
    var sel = sh.getActiveRange(); var r = sel.getRow();
    var head = headers_(sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0]);

    var levelId = (sh.getRange(r, head.LevelID+1).getValue() || '').toString().trim();
    var url     = (sh.getRange(r, head.ShareURL+1).getValue() || '').toString().trim();
    if (!url || !levelId) { SpreadsheetApp.getUi().alert('Select a row with LevelID and ShareURL.'); return; }

    var ch = extractChannelId_(url);
    var fetched = ch ? fetchGameLabSource_(ch) : {ok:false, msg:'No ChannelID'};
    var src = fetched && fetched.ok ? fetched.src : '';
    var snippet = (src || '').slice(0, 300).replace(/\n/g, '\\n');

    var critByLevel = loadCriteriaByLevel_();
    var crits = critByLevel[levelId] || [];

    var msg =
      'LevelID: ' + levelId + '\n' +
      'ChannelID: ' + (ch || '(none)') + '\n' +
      'Fetch: ' + (fetched && fetched.ok ? 'OK' : ('FAIL: ' + (fetched && fetched.msg))) + '\n' +
      'Code length: ' + (src ? src.length : 0) + '\n' +
      'Criteria count: ' + crits.length + '\n\n' +
      'First 300 chars:\n' + snippet;
    SpreadsheetApp.getUi().alert(msg);
  } catch (e) {
    SpreadsheetApp.getUi().alert('Debug failed: ' + e);
  }
}

function isAutoGrade_(){ 
  var v = PropertiesService.getScriptProperties().getProperty('AUTO_GRADE'); 
  return v == null ? true : String(v).toLowerCase() !== 'false';
}
function isAutoEmail_(){ 
  var v = PropertiesService.getScriptProperties().getProperty('AUTO_EMAIL'); 
  return v == null ? true : String(v).toLowerCase() !== 'false';
}
function setFlag_(k, val){ 
  PropertiesService.getScriptProperties().setProperty(k, val ? 'true' : 'false'); 
}
function toggleAutoGrade(){ setFlag_('AUTO_GRADE', !isAutoGrade_()); onOpen(); SpreadsheetApp.getUi().alert('Auto-grade is now ' + (isAutoGrade_()?'ON':'OFF')); }
function toggleAutoEmail(){ setFlag_('AUTO_EMAIL', !isAutoEmail_()); onOpen(); SpreadsheetApp.getUi().alert('Auto-email is now ' + (isAutoEmail_()?'ON':'OFF')); }

function headersSmart_(row1) {
  var aliases = {
    Timestamp:   ['Timestamp','Response Timestamp','Submitted at'],
    Email:       ['Email','Email Address','Email address'],
    First:       ['First','First Name','Given Name'],
    Last:        ['Last','Last Name','Family Name','Surname'],
    Period:      ['Period','Class Period','Class','Section'],
    LevelID:     ['LevelID','Level ID','Which assessment level','Which assessment level are you submitting'],
    ShareURL:    ['ShareURL','Share URL','URL','Project URL','Project Link','Paste the URL to your completed assessment level'],
  };
  var map = {};
  for (var c = 0; c < row1.length; c++) {
    var h = String(row1[c] || '').trim();
    var hl = h.toLowerCase();
    Object.keys(aliases).forEach(function(key){
      if (map[key] !== undefined) return;
      aliases[key].some(function(alias){
        var al = alias.toLowerCase();
        if (hl === al || hl.indexOf(al) === 0) { map[key] = c; return true; }
        return false;
      });
    });
  }
  return map;
}

function backfillFormResponses() {
  backfillFormResponsesCore_(false);
}
function backfillFormResponsesAndEmail() {
  backfillFormResponsesCore_(true);
}

function backfillFormResponsesCore_(sendEmails) {
  var ss = SpreadsheetApp.getActive();
  var subSh = getSheet_(SHEET_SUB); // "Submissions"
  var subHeaders = subSh.getRange(1,1,1,subSh.getLastColumn()).getValues()[0];
  var subMap = headers_(subHeaders);

  // Find the source sheet (Form writes here)
  var srcSh = ss.getSheetByName('Form Responses 1') || findFormResponsesSheet_();
  if (!srcSh) { SpreadsheetApp.getUi().alert('No "Form Responses" sheet found.'); return; }

  var srcValues = srcSh.getDataRange().getValues();
  if (srcValues.length <= 1) { SpreadsheetApp.getUi().alert('No rows to backfill.'); return; }
  var srcHead = srcValues[0];
  var srcMap = headersSmart_(srcHead); // maps verbose question text to our short names

  // Build existing key set from Submissions: key = Timestamp + Email + LevelID
  var existing = {};
  var subValues = subSh.getDataRange().getValues();
  for (var i=1; i<subValues.length; i++) {
    var ts = normalizeTimestamp_(subValues[i][subMap.Timestamp]);
    var em = (subMap.Email !== undefined) ? String(subValues[i][subMap.Email] || '').trim().toLowerCase() : '';
    var lvl = (subMap.LevelID !== undefined) ? String(subValues[i][subMap.LevelID] || '').trim() : '';
    var key = [ts, em, lvl].join('|');
    existing[key] = true;
  }

  var appendedRows = [];
  for (var r=1; r<srcValues.length; r++) {
    var row = srcValues[r];

    var tsVal = (srcMap.Timestamp !== undefined) ? row[srcMap.Timestamp] : new Date();
    var emVal = (srcMap.Email     !== undefined) ? row[srcMap.Email]     : '';
    var first = (srcMap.First     !== undefined) ? row[srcMap.First]     : '';
    var last  = (srcMap.Last      !== undefined) ? row[srcMap.Last]      : '';
    var period= (srcMap.Period    !== undefined) ? row[srcMap.Period]    : '';
    var level = (srcMap.LevelID   !== undefined) ? row[srcMap.LevelID]   : '';
    var share = (srcMap.ShareURL  !== undefined) ? row[srcMap.ShareURL]  : '';

    var key = [normalizeTimestamp_(tsVal), String(emVal||'').trim().toLowerCase(), String(level||'').trim()].join('|');
    if (existing[key]) continue; // skip duplicates

    // Construct a normalized Submissions row
    var out = new Array(subHeaders.length).fill('');
    if (subMap.Timestamp !== undefined) out[subMap.Timestamp] = tsVal || new Date();
    if (subMap.Email     !== undefined) out[subMap.Email]     = emVal;
    if (subMap.First     !== undefined) out[subMap.First]     = first;
    if (subMap.Last      !== undefined) out[subMap.Last]      = last;
    if (subMap.Period    !== undefined) out[subMap.Period]    = period;
    if (subMap.LevelID   !== undefined) out[subMap.LevelID]   = level;
    if (subMap.ShareURL  !== undefined) out[subMap.ShareURL]  = share;

    subSh.appendRow(out);
    var newRow = subSh.getLastRow();
    appendedRows.push(newRow);
    existing[key] = true;
  }

  if (appendedRows.length) {
    gradeRows_(appendedRows);
    if (sendEmails) {
      appendedRows.forEach(function(r){ try { sendEmailForRow_(r); } catch(_) {} });
    }
  }

  SpreadsheetApp.getUi().alert(
    'Backfill complete: appended ' + appendedRows.length + ' new row(s) from "' + srcSh.getName() + '".'
  );
}

// Normalize timestamp to a stable string
function normalizeTimestamp_(v) {
  try {
    var tz = Session.getScriptTimeZone() || 'UTC';
    var d = (v instanceof Date) ? v : new Date(v);
    if (isNaN(d.getTime())) return String(v||'').trim();
    return Utilities.formatDate(d, tz, "yyyy-MM-dd'T'HH:mm:ss");
  } catch(e) {
    return String(v||'').trim();
  }
}

// Find a sheet named like "Form Responses", e.g., "Form Responses 1"
function findFormResponsesSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('Form Responses 1') || ss.getSheetByName('Form Responses');
  if (sh) return sh;
  var all = ss.getSheets();
  for (var i=0;i<all.length;i++) {
    if (/^Form Responses/i.test(all[i].getName())) return all[i];
  }
  return null;
}

