const test = () => {
  const correctResult = getIssuesAndSuggestionsForAdText('What goes into a great books text? How can you write a text that drives people to click through and convert?');
  if (correctResult !== 'correct!') {
    throw Error('Failed test with correct case.');
  }

  const mispelledResult = getIssuesAndSuggestionsForAdText('What goes into a great books text? Howcan you write a text that drives people to click through and convert?');
  if (mispelledResult === 'correct!' && mispelledResult === 'error') {
    throw Error('Failed test with mispelled case.');
  }
};

const getIssuesAndSuggestionsForAdText = (adText) => {
  try {
    const bing = new BingSpellChecker({
      key : 'f73d4ac33bbd441491a096c8eec150a5',
      toIgnore : [],
      enableCache : false
    });

    const issues = bing.getSpellingIssues(adText);
    bing.saveCache();

    if (issues.length > 0) {
      const issueAndFirstSuggestions = [];
      for (const issue of issues) {
        issueAndFirstSuggestions.push(issue.token + ' :: ' + issue.suggestions[0]["suggestion"]);
      }

      return issueAndFirstSuggestions.join('\n');
    }

    return 'correct!';
  } catch(e) {
    Logger.log('INFO: '+e);

    return 'error';
  }
};

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

  if (config.enableCache) {
    this.loadCache();
  }

  // Given a set of options, this function calls the API to check the spelling
  // options:
  //   options.text : the text to check
  //   options.mode : the mode to use, defaults to 'proof'
  // returns a list of misspelled words, or empty list if everything is good.
  this.checkSpelling = (options) => {
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
  this.getSpellingIssues = (toCheck) => {
    const issues = this.checkSpelling({ text : toCheck });
    if (issues.length > 0) {
      Logger.log('Checked text: %s \n Issues found: %s', toCheck, JSON.stringify(issues));
    }
    return issues;
  };

  // Loads the list of misspelled words from Google Drive.
  // set config.enableCache to true to enable.
  this.loadCache = () => {
    const fileIter = DriveApp.getFilesByName(this.CACHE_FILE_NAME);
    if(fileIter.hasNext()) {
      this.cache = JSON.parse(fileIter.next().getBlob().getDataAsString());
    } else {
      this.cache = { incorrect : {} };
    }
  };

  // Called when you are finished with everything to store the data back to Google Drive
  this.saveCache = () => {
    const fileIter = DriveApp.getFilesByName(this.CACHE_FILE_NAME);
    if(fileIter.hasNext()) {
      fileIter.next().setContent(JSON.stringify(this.cache));
    } else {
      DriveApp.createFile(this.CACHE_FILE_NAME, JSON.stringify(this.cache));
    }
  };
}