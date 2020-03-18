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
class BingSpellChecker {
  config = {};
  BASE_URL = 'https://api.cognitive.microsoft.com/bing/v7.0/spellcheck';
  CACHE_FILE_NAME = 'spellcheck_cache.json';
  key = config.key;
  toIgnore = config.toIgnore;
  cache = null;
  previousText = null;
  previousResult = null;
  delay = (config.delay) ? config.delay : 60000/7;
  timeOfLastCall = null;
  hitQuota = false;

  constructor(configP) {
    this.config = configP;
    if (this.config.enableCache) {
      this.loadCache();
    }
  }

  // Given a set of options, this function calls the API to check the spelling
  // options:
  //   options.text : the text to check
  //   options.mode : the mode to use, defaults to 'proof'
  // returns a list of misspelled words, or empty list if everything is good.
  checkSpelling = (options) => {
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
        const words = options.text.split(/ +/);
        for(const i in words) {
          //Logger.log('INFO: checking cache: '+words[i]);
          if(this.cache && this.cache.incorrect[words[i]]) {
            //Logger.log('INFO: Using cached response.');
            return [{"offset":1,"token":words[i],"type":"cacheHit","suggestions":[]}];
          }
        }
      }

      let url = this.BASE_URL;
      const config = {
        method : 'POST',
        headers : {
          'Ocp-Apim-Subscription-Key' : this.key,
          'Content-Type' : 'application/x-www-form-urlencoded'
        },
        payload : 'Text='+encodeURIComponent(options.text),
        muteHttpExceptions : true
      };
      if (options && options.mode) {
        url += '?mode='+options.mode;
      } else {
        url += '?mode=proof';
      }

      if (this.timeOfLastCall) {
        const now = Date.now();
        if(now - this.timeOfLastCall < this.delay) {
          // Logger.log(Utilities.formatString('INFO: Sleeping for %s milliseconds',
          //     this.delay - (now - this.timeOfLastCall)));
          Utilities.sleep(this.delay - (now - this.timeOfLastCall));
        }
      }

      const resp = UrlFetchApp.fetch(url, config);
      this.timeOfLastCall = Date.now();

      if(resp.getResponseCode() !== 200) {
        if(resp.getResponseCode() === 403) {
          this.hitQuota = true;
        }
        throw JSON.parse(resp.getContentText()).message;
      } else {
        const jsonResp = JSON.parse(resp.getContentText());
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
  getSpellingIssues = (toCheck) => {
    const issues = this.checkSpelling({ text : toCheck });
    if (issues.length > 0) {
      Logger.log('Checked text: %s \n Issues found: %s', toCheck, JSON.stringify(issues));
    }
    return issues;
  };

  // Loads the list of misspelled words from Google Drive.
  // set config.enableCache to true to enable.
  loadCache = () => {
    const fileIter = DriveApp.getFilesByName(this.CACHE_FILE_NAME);
    if(fileIter.hasNext()) {
      this.cache = JSON.parse(fileIter.next().getBlob().getDataAsString());
    } else {
      this.cache = { incorrect : {} };
    }
  };

  // Called when you are finished with everything to store the data back to Google Drive
  saveCache = () => {
    const fileIter = DriveApp.getFilesByName(this.CACHE_FILE_NAME);
    if(fileIter.hasNext()) {
      fileIter.next().setContent(JSON.stringify(this.cache));
    } else {
      DriveApp.createFile(this.CACHE_FILE_NAME, JSON.stringify(this.cache));
    }
  };
}

module.exports = {
  BingSpellChecker: BingSpellChecker
};