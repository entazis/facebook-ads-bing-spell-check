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
    const BingSpellChecker = require('BingSpellChecker');
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