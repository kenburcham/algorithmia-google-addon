/**
* Algorithmia Google Sheets Add-on
* Ken Burcham - January 2018
*
* Provides a sidebar, client and ALGO function to enable the use of Algorithmia
* "serverless" AI/utility functions in your Google Sheets. 
*/


/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */


//load our apikey from user properties
var api_key = getUserPropertyValue('API_KEY');

if(!api_key)
{
  Logger.log("No api key found.");
}
  

//couple of constants
var DIRECTION_RIGHT = 1;
var DIRECTION_DOWN = 2;

/* What should the add-on do after it is installed */
function onInstall(e) {
  onOpen(e);
}

/* What should the add-on do when a document is opened */
function onOpen(e) {  
  var ui = SpreadsheetApp.getUi();  // Or DocumentApp or FormApp.
  ui.createMenu('Algorithmia')
      .addItem('Use Algorithms', 'showSidebar')
      .addItem('Configuration','showConfig')
  .addToUi();
}

/* Show the Settings in the sidebar */
function showConfig(){
  var html = HtmlService.createTemplateFromFile("Config")
    .evaluate()
    .setTitle("Algorithmia config"); // The title shows in the sidebar
  
  SpreadsheetApp.getUi().showSidebar(html); 
  
}
 
/**
* Gets the given user property from the PropertiesService
* @param {string} key The property key to look up.
* @returns {string} The value returned by the PropertyService for this key.
*/
function getUserPropertyValue(key){
  //Logger.log("getUserPropertyValue called - " + key);
  var userProperties = PropertiesService.getUserProperties();
  var we_got = userProperties.getProperty(key);
  //Logger.log("we got: ");
  //Logger.log(we_got);
  return we_got;
}



/**
* Sets the give user property in the PropertiesService
* @param {string} key The property key to set.
* @param {string} value The property value to set.
* @returns "failed" or "saved"
*/
function setUserPropertyValue(key, value){
   // Logger.log("setUserPropertyValue called - " + key + " - " + value);
  try{
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty(key, value);
    return "saved."
  }catch(e){
    Logger.log(e);
    return "failed.";
  }
}

/* Show the main sidebar */
function showSidebar() {
  var html = HtmlService.createTemplateFromFile("Sidebar")
    .evaluate()
    .setTitle("Algorithmia"); // The title shows in the sidebar
  SpreadsheetApp.getUi().showSidebar(html);
}


/**
* Call from a cell in a Google Sheet to make a call to a Algorithmia algorithm API.
* Note: the algorithm must be already defined below.
* @param {string} a_algorithm The algorithm that you want to run. must match a configured algorithm in the array below
* @param {string} a_input A cell or value that you want to provide as input
* @param {string} a_options Any options that you want to send to your algorithm function defined below
* @returns result of the request to put into the cell
*/
function ALGO(a_algorithm, a_input, a_options) {
  if(!a_algorithm || !a_input)
  {
    alert("ALGO needs an algorithm with input in order to run.");
    return;
  }
  
  Logger.log("running ALGO for " + a_algorithm + " on ( " + a_input + " )");
  
  var algo_to_run = algorithms[a_algorithm].action;
  return algo_to_run(a_input, a_options);
}

algorithms = {};


//add an algorithm to our library
function add_algorithm(a_algorithm){
  algorithms[a_algorithm.name]=a_algorithm;
}



/** 
* Called from the sidebar to run an algorithm 
* @param {string} a_algorithm Name of the algorithm to run
* @returns void
*/
function runAlgorithm(a_algorithm, a_algoURL){
  Logger.log("runAlgorithm: " + a_algorithm);
  Logger.log("algoURL: " + a_algoURL);
  
  sheet = SpreadsheetApp.getActiveSheet();
  
  a_cell = sheet.getActiveCell();
  Logger.log("we are on: " + a_cell.getA1Notation());
  
  algorithm = getAlgorithm(a_algorithm);
  
  //if this algo is a formula then just write it to the adjacent cell.
  if(algorithm.isFormula){
  
     a_cell_adjacent = sheet.getRange(a_cell.getRow(), a_cell.getColumn()+1);
     //set the cell to the right to be: =ALGO("my-algorithm",C7)
     a_cell_adjacent.setValue('=ALGO("'+a_algorithm+'",'+a_cell.getA1Notation()+')');
     sheet.setActiveRange(a_cell_adjacent);
  }
  else //the action will write out on its own...
  {
     a_cell_adjacent = sheet.getRange(a_cell.getRow(), a_cell.getColumn()+1);
     //set the cell to the right to be: =ALGO("my-algorithm",C7)
     a_cell_adjacent.setValue("Executing...");
    
     Logger.log("running non-formula from sidebar for " + a_algorithm );
     var algo_to_run = algorithm.action;
    
     if(a_algorithm == 'user-defined')
       algo_to_run(a_cell.getValue(), a_algoURL);
     else
       algo_to_run(a_cell.getValue());
  }
}


/**
* Called by the sidebar to populate the dropdown of possible algorithms
* @returns all configured algorithms
*/
function getAlgorithms(){
  return algorithms; 
}

function getAlgorithm(name)
{
  Logger.log("getting an algorithm: " + name);
  return algorithms[name]; 
}

/**
* each call to add_algorithm below configures a new algorithm for use.
* 
* to add a new algorithm, just copy this pattern but be sure to setup the input and the output
* according to the needs of the algorithm defined at Algorithmia.com
*
* note that these are simply JSON objects and can later be retrieved from somewhere else and iterated.
*/
add_algorithm(
  {
    name:"analyze-url",
    label: "Analyze URL: Summary",
    isFormula: true, //when this is true, our call will return a single value so we can use the ALGO formula
    description: "Analyze a URL and return a summary of the page in the next cell. <hr/>=ALGO('analyze-sentiment')<hr/>Inserts ALGO function in the adjacent cell to the right of selected cell.",
    action: function(a_inputURL){

       var input = [a_inputURL];

       if(input == "")
         throw new Error("AnalyzeURL requires an input URL");

       result = Algorithmia.client(api_key)
           .algo("algo://web/AnalyzeURL/0.2.17")
           .pipe(input);

       // in our case, we want the "summary" field from the resulting JSON
       //Logger.log(result);
       if(result.summary === "")
         return "<no summary>";
       else
         return result.summary;
     },
  }
);


add_algorithm(
  {
    name: "analyze-sentiment",
    label: "Analyze Sentiment",
    isFormula: true,
    description: "Analyze a cell's text contents and return a sentiment score from -1 (very negative) to +1 (very positive) where 0 is neutral. <hr/>=ALGO('analyze-sentiment')<hr/>Inserts ALGO function in the adjacent cell to the right of selected cell.",
    action: function(a_inputText){
      var input = {"document": a_inputText};
      Logger.log("running analyze sentiment on "+a_inputText);
      Logger.log(input);
      
      var result = Algorithmia.client(api_key)
         .algo("nlp/SentimentAnalysis/1.0.4")
         .pipe(input);
      
      Logger.log("results = ");
      Logger.log(result);
      
      if(result[0].sentiment === "")
        return "<no sentiment>";
      else
        return result[0].sentiment;
      
    }
  });

//remember whenever adding a new algorithm, the INPUT needs to be prepped and the OUTPUT read for the right result.
add_algorithm(
  {
    name: "email-validator",
    label: "Email Validator",
    isFormula: true,
    description: "Validate an email address meets just the most basic structure.<hr/>=ALGO('email-validator')<hr/>Inserts ALGO function in the adjacent cell to the right of selected cell.",
    action: function(a_inputText){
      var input = a_inputText;
      
      Logger.log("running avlidate email.");
      Logger.log(input);
      
      var result = Algorithmia.client(api_key)
         .algo("diego/EmailValidator/0.1.0")
         .pipe(input);
      
      return result;
      
    }
  });



add_algorithm(
  {
    name: "social-media-shares",
    label: "Social Media Shares",
    isFormula: false, //because we return more than one results
    description: "For a web address (URL) find the number of social media shares for Facebook, Pinterest and LinkedIn. (20 credit royalty). <hr/>Runs algorithm and copies results to the adjacent cells to the right.",
    action: function(a_inputText){
      var input = a_inputText;
      
      if(!a_inputText){
        Logger.log("No URL given for social-media-shares algorithm to check.");
        return;
      }
      
      var result = Algorithmia.client(api_key)
         .algo("web/ShareCounts/0.2.8")
         .pipe(input);
      
      Logger.log(result);
      
      if(result){
        copyToAdjacentCells(
          [
            result.facebook_likes,
            result.facebook_shares,
            result.facebook_comments,
            result.pinterest
          ]);
      }else{
        copyToAdjacentCells(["no result"]);
      }
    }
  });


//brunni/AddressExtraction/0.1.2
add_algorithm(
  {
    name: "contact-extraction",
    label: "Contact Info Extraction",
    isFormula: false, //because we return more than one results
    description: "For a web address (URL) extract the contact information (100 credit royalty). <hr/>Runs algorithm and copies results to adjacent cells to the right.",
    action: function(a_inputText){
      var input = a_inputText;
      
      if(!a_inputText){
        Logger.log("No URL given for algorithm to check.");
        return;
      }
      
      Logger.log("running contact extraction on " + input);
      
      var result = Algorithmia.client(api_key)
         .algo("brunni/AddressExtraction/0.1.2")
         .pipe(input);
      
      Logger.log(result);
      
      if(result && result.results.length > 0){
        copyToAdjacentCells(
          [
            result.results[0].company,
            result.results[0].street,
            result.results[0].city,
            result.results[0].zip,
            result.results[0].country,
            result.results[0].managers,
            result.social_links,
            result.url_imprint
          ]);
      }else{
        copyToAdjacentCells(["no result"]);
      }
      
    }
  });

//web/GetLinks/0.1.5
add_algorithm(
  {
    name: "get-links",
    label: "Extract Links from Page",
    isFormula: false, //because we return more than one results
    description: "For a webpage, extract all of the URLs on the page into a list. <hr/>Runs algorithm and copies results DOWN the cells in the adjacent column to the right",
    action: function(a_inputText){
      var input = a_inputText;
      
      if(!a_inputText){
        Logger.log("No URL given for algorithm to check.");
        return;
      }
      
      Logger.log("running url extraction on " + input);
      
      var result = Algorithmia.client(api_key)
         .algo("web/GetLinks/0.1.5")
         .pipe(input);
      
      Logger.log(result);
      
      if(result && result.length > 0){
        copyToAdjacentCells(result, DIRECTION_DOWN);
      }else{
        copyToAdjacentCells(["no result"]);
      }
      
    }
  });


//jhurliman/GetEMailAddresses/0.1.2
add_algorithm(
  {
    name: "get-emails",
    label: "Extract Emails from Page",
    isFormula: false, //because we return more than one results
    description: "For a webpage, extract all of the email addresses on the page into a list (10 credit royalty).<hr/>Runs algorithm and copies results DOWN the cells in the adjacent column to the right.",
    action: function(a_inputText){
      var input = { url: a_inputText };
      
      if(!a_inputText){
        Logger.log("No URL given for algorithm to check.");
        return;
      }
      
      Logger.log("running email extraction on " + input);
      
      var result = Algorithmia.client(api_key)
         .algo("jhurliman/GetEMailAddresses/0.1.2")
         .pipe(input);
      
      Logger.log(result);
      
      if(result && result.emails && result.emails.length > 0){
        copyToAdjacentCells(result.emails, DIRECTION_DOWN);
      }else{
        copyToAdjacentCells(["no result"]);
      }
      
    }
  });

//diego/RetrieveTweetsWithKeyword/0.1.2
add_algorithm(
  {
    name: "tweet-feeder",
    label: "Search Twitter",
    isFormula: false, //because we return more than one results
    description: "Search Twitter for matches to a keyword. <hr/>Runs algorithm and copies results DOWN the cells in the adjacent column to the right.",
    action: function(a_inputText){
      var input = a_inputText;
      
      if(!a_inputText){
        Logger.log("No URL given for algorithm to check.");
        return;
      }
      
      Logger.log("running twitter search on " + input);
      
      var result = Algorithmia.client(api_key)
         .algo("diego/RetrieveTweetsWithKeyword/0.1.2")
         .pipe(input);
      
      Logger.log(result);
      
      if(result && result.length > 0){
        copyToAdjacentCells(result, DIRECTION_DOWN);
      }else{
        copyToAdjacentCells(["no result"]);
      }
      
    }
  });

add_algorithm(
  {
    name: "user-defined",
    label: "User Defined",
    isFormula: false,
    showALGOfield: true,
    description: "<b>Experimental!</b> Run your own choice of algorithm by pasting the URL from algorithmia into the field above. <br/> 1) Browse <a href='https://algorithmia.com/algorithms' target='_blank'>Algorithmia Marketplace</a> and find an algorithm you want to run on your data. <br/> 2) Paste the algorithm address (like: 'nlp/Summarizer/0.1.6' into the field above. <br/> 3) Click GO!<hr/>Your selected cell's value will be sent as the input and the output will be displayed in the adjacent field to the right. (Maybe. If it works!)",
    action: function(a_inputText, a_algoURL){
      var input = a_inputText;
      
      if(!a_inputText){
        Logger.log("No input given for algorithm");
        return;
      }
      
      Logger.log("running user-defined algorithm " + a_algoURL + " on input: " + input);
      
      var result = Algorithmia.client(api_key)
         .algo(a_algoURL)
         .pipe(input);
      
      Logger.log(result);
      
      if(result)
      {
        if(Array.isArray(result)){
         copyToAdjacentCells(result, DIRECTION_DOWN);
        }else{
          copyToAdjacentCells([result]);
        }
      }else{
        copyToAdjacentCells(["no result"]);
      }
    }
  });
      
/**
* takes an array of values and writes them out to the adjacent cells
* @param {Array} a_data An array of data to copy into the spreadsheet.
* @param {number} a_direction One of the constants defined above that indicates direction to copy array values. defaults to DIRECTION_RIGHT
* @returns void
*/
function copyToAdjacentCells(a_data, a_direction) {

  if(!a_direction)
    a_direction = DIRECTION_RIGHT; 
  
  if(!a_data || !Array.isArray(a_data))
  {
    Logger.log("something went wrong -- a_data isn't an array");
    Logger.log(a_data);
    return;
  }
  
  sheet = SpreadsheetApp.getActiveSheet();
  
  a_cell = sheet.getActiveCell();
  Logger.log("we are on: " + a_cell.getA1Notation());

  //target is the next adjacent cell...
  var target_row = a_cell.getRow();
  var target_col = a_cell.getColumn()+1;
  
  Logger.log("adjacent is " + target_row + " - " + target_col);
  
  //jump to adjacent cell.
  sheet.setActiveRange(sheet.getRange(target_row, target_col));
  
  //copy the values to all the next adjacent rows.
  a_data.forEach(function(item){
    cell_adjacent = sheet.getRange(target_row, target_col);
    cell_adjacent.setValue(item);
    
    //increment whichever direction we're going...
    if(a_direction === DIRECTION_RIGHT)
      target_col ++;
    else
      target_row ++;
    
  });
  
  
  
  //label the headers
  //for(key in results) {
  //  target_col++;
  //  sheet.getRange(target_row, target_col).setValue(key); //key = label
  //}

  //target_col = active_col; //reset to beginning column
  //target_row++; //move to the next row...

  //copy each json field in the top level to each consecutive cell to the right AFTER the current one
  //for(key in results)  {
  //  target_col++;
  //  sheet.getRange(target_row, target_col).setValue(results[key].toString().substring(0,255));
  //}

}