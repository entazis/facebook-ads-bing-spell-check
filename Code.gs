var CHECKED_LABEL_NAME = "Spellchecked";
var ISSUE_LABEL_NAME = "Spelling Issue";

function main() {
  var bing = new BingSpellChecker({
    key : '',
    toIgnore : ['adwords','adgroup','russ'],
    enableCache : false
  });
  //Optional parameters for filtering account names.
  //Leave blank to use filters. The matching is case insensitive.
  var excludeAccountNameContains = "- OLD"; //Select which accounts to exclude. Leave blank to not exclude any accounts.
  var includeAccountNameContains = ""; //Select which accounts to include. Leave blank to include all accounts.

  var emailAddressesForSendingReportTo = [''];

  var sheet = openSpreadsheetAndGetSheet('', 'Spelling Issues');
  Logger.log("Using sheet: %s", sheet.getName());

  var accountIter = MccApp.accounts()
    .withCondition('Name DOES_NOT_CONTAIN_IGNORE_CASE "' + excludeAccountNameContains + '"')
    .withCondition('Name CONTAINS_IGNORE_CASE "' + includeAccountNameContains + '"')
    .get();
  while(accountIter.hasNext()) {
    MccApp.select(accountIter.next());
    checkAds(bing);
    if(bing.hitQuota) {
      break;
    }
  }

  sendSpellCheckReportEmail(emailAddressesForSendingReportTo);

  bing.saveCache();
}

function checkAds(bing) {
  createLabelIfNeeded(CHECKED_LABEL_NAME,'Indicates an entity was spell checked','#00ff00' /*green*/);
  createLabelIfNeeded(ISSUE_LABEL_NAME,'Indicates an entity has a spelling issue','#ff0000' /*red*/);

  var adIter = AdWordsApp.ads()
    .withCondition("Status = ENABLED")
    .withCondition(Utilities.formatString("LabelNames CONTAINS_NONE ['%s','%s']",
      CHECKED_LABEL_NAME,
      ISSUE_LABEL_NAME))
    .get();
  while(adIter.hasNext() && !bing.hitQuota) {
    var ad = adIter.next();
    var textToCheck = "";
    if (ad.getType() === "EXPANDED_TEXT_AD") {
      var expandedTextAd = ad.asType().expandedTextAd();
      textToCheck = [
        expandedTextAd.getHeadlinePart1(),
        expandedTextAd.getHeadlinePart2(),
        expandedTextAd.getDescription()
      ].join(' ');
    } else {
      textToCheck = [
        ad.getHeadline(),
        ad.getDescription1(),
        ad.getDescription2()
      ].join(' ');
    }
    try {
      var issues = bing.getSpellingIssues(textToCheck);
      if(issues.length > 0) {
        ad.applyLabel(ISSUE_LABEL_NAME);
        var missSpellings = [];
        var suggestions = [];

        for (var i = 0; i < issues.length; i++) {
          missSpellings.push(issues[i].token);
          suggestions.push(issues[i].suggestions[0]["suggestion"]);
        }

        appendARow(AdWordsApp.currentAccount().getName(), ad.getCampaign().getName(), ad.getAdGroup().getName(), ad.getId(), missSpellings.join('\n'), suggestions.join('\n'));
      } else {
        ad.applyLabel(CHECKED_LABEL_NAME);
      }
    } catch(e) {
      // This probably means you're out of quota.
      // You can pick up from here next time.
      Logger.log('INFO: '+e);
      break;
    }
    if(!AdWordsApp.getExecutionInfo().isPreview() &&
      AdWordsApp.getExecutionInfo().getRemainingTime() < 60) {
      // Out of time
      Logger.log("INFO: Ran out of time. Will continue next run.");
      break;
    }
  }
}

//This is a helper function to create the label if it does not already exist
function createLabelIfNeeded(name,description,color) {
  if(!AdWordsApp.labels().withCondition("Name = '"+name+"'").get().hasNext()) {
    AdWordsApp.createLabel(name,description,color);
  }
}

/******************************************
 * Bing Spellchecker API v1.0
 * By: Russ Savage (@russellsavage)
 * Usage:
 *  // You will need a key from
 *  // https://www.microsoft.com/cognitive-services/en-us/bing-spell-check-api/documentation
 *  // to use this library.
 *  var bing = new BingSpellChecker({
 *    key : 'xxxxxxxxxxxxxxxxxxxxxxxxx',
 *    toIgnore : ['list','of','words','to','ignore'],
 *    enableCache : true // <- stores data in a file to reduce api calls
 *  });
 * // Example usage:
 * var hasSpellingIssues = bing.hasSpellingIssues('this is a speling error');
 ******************************************/
function BingSpellChecker(config) {
  this.BASE_URL = 'https://api.cognitive.microsoft.com/bing/v7.0/spellcheck';
  this.CACHE_FILE_NAME = 'spellcheck_cache.json';
  this.key = config.key;
  this.toIgnore = config.toIgnore;
  this.cache = null;
  this.previousText = null;
  this.previousResult = null;
  this.delay = (config.delay) ? config.delay : 60000/7;
  this.timeOfLastCall = null;
  this.hitQuota = false;

  // Given a set of options, this function calls the API to check the spelling
  // options:
  //   options.text : the text to check
  //   options.mode : the mode to use, defaults to 'proof'
  // returns a list of misspelled words, or empty list if everything is good.
  this.checkSpelling = function(options) {
    if(this.toIgnore) {
      options.text = options.text.replace(new RegExp(this.toIgnore.join('|'),'gi'), '');
    }
    options.text = options.text.replace(/{.+}/gi, '');
    options.text = options.text.replace(/[^a-z ]/gi, '').trim();

    if(options.text.trim()) {
      if(options.text === this.previousText) {
        //Logger.log('INFO: Using previous response.');
        return this.previousResult;
      }
      if(this.cache) {
        var words = options.text.split(/ +/);
        for(const i in words) {
          //Logger.log('INFO: checking cache: '+words[i]);
          if(this.cache && this.cache.incorrect[words[i]]) {
            //Logger.log('INFO: Using cached response.');
            return [{"offset":1,"token":words[i],"type":"cacheHit","suggestions":[]}];
          }
        }
      }
      var url = this.BASE_URL;
      var config = {
        method : 'POST',
        headers : {
          'Ocp-Apim-Subscription-Key' : this.key,
          'Content-Type' : 'application/x-www-form-urlencoded'
        },
        payload : 'Text='+encodeURIComponent(options.text),
        muteHttpExceptions : true
      };
      if(options && options.mode) {
        url += '?mode='+options.mode;
      } else {
        url += '?mode=proof';
      }
      if(this.timeOfLastCall) {
        var now = Date.now();
        if(now - this.timeOfLastCall < this.delay) {
          // Logger.log(Utilities.formatString('INFO: Sleeping for %s milliseconds',
          //     this.delay - (now - this.timeOfLastCall)));
          Utilities.sleep(this.delay - (now - this.timeOfLastCall));
        }
      }
      var resp = UrlFetchApp.fetch(url, config);
      this.timeOfLastCall = Date.now();
      if(resp.getResponseCode() !== 200) {
        if(resp.getResponseCode() === 403) {
          this.hitQuota = true;
        }
        throw JSON.parse(resp.getContentText()).message;
      } else {
        var jsonResp = JSON.parse(resp.getContentText());
        this.previousText = options.text;
        this.previousResult = jsonResp.flaggedTokens;
        for(const i in jsonResp.flaggedTokens) {
          if (this.cache) {
            this.cache.incorrect[jsonResp.flaggedTokens[i].token] = true;
          }
        }
        return jsonResp.flaggedTokens;
      }
    } else {
      return [];
    }
  };

  // Returns the spelling issues if there are spelling mistakes in the text toCheck
  // toCheck : the phrase to spellcheck
  // returns array of objects if there are words misspelled, empty array otherwise.
  this.getSpellingIssues = function(toCheck) {
    var issues = this.checkSpelling({ text : toCheck });
    if (issues.length > 0) {
      Logger.log('Checked text: %s \n Issues found: %s', toCheck, JSON.stringify(issues));
    }
    return issues;
  };

  // Loads the list of misspelled words from Google Drive.
  // set config.enableCache to true to enable.
  this.loadCache = function() {
    var fileIter = DriveApp.getFilesByName(this.CACHE_FILE_NAME);
    if(fileIter.hasNext()) {
      this.cache = JSON.parse(fileIter.next().getBlob().getDataAsString());
    } else {
      this.cache = { incorrect : {} };
    }
  };

  if(config.enableCache) {
    this.loadCache();
  }

  // Called when you are finished with everything to store the data back to Google Drive
  this.saveCache = function() {
    var fileIter = DriveApp.getFilesByName(this.CACHE_FILE_NAME);
    if(fileIter.hasNext()) {
      fileIter.next().setContent(JSON.stringify(this.cache));
    } else {
      DriveApp.createFile(this.CACHE_FILE_NAME, JSON.stringify(this.cache));
    }
  }
}

function openSpreadsheetAndGetSheet(url, sheetName) {
  var ss = SpreadsheetApp.openByUrl(url);
  return ss.getSheetByName(sheetName);
}

function appendARow(accountName, campaignName, adGroupName, adId, issues, suggestions) {
  var sheet = openSpreadsheetAndGetSheet('', 'Spelling Issues');
  sheet.appendRow([accountName, campaignName, adGroupName, adId, issues, suggestions]);
}

function sendSpellCheckReportEmail(emails) {
  var sheet = openSpreadsheetAndGetSheet('', 'Spelling Issues');
  var uniqueIssuesCnt = countDistinctValues(sheet.getSheetValues(2, 5, sheet.getLastRow(), 1));

  for (var i = 0; i < emails.length; i++) {
    MailApp.sendEmail(emails[i],
      'Bing Spell Checker - report',
      Utilities.formatString('Spell check was successful!\n\nSpelling issues were logged here: https://docs.google.com/spreadsheets/d/1r3B_sZhPySjd1GRnIooeRSmj8Bg6TJ5SDN0J0XwW_uE/\nNumber of unique issues: %s', uniqueIssuesCnt));
  }
}

function countDistinctValues(values) {
  values = values.filter(function(value) {
    return !(JSON.stringify(value) === '[""]');
  });

  var counts = {};
  for (var i = 0; i < values.length; i++) {
    counts[values[i]] = 1 + (counts[values[i]] || 0);
  }

  return Object.keys(counts).length;
}