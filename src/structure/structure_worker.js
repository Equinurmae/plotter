importScripts('https://unpkg.com/compromise@latest/builds/compromise.min.js');

// returns the word count
self.onmessage = function(e) {
    self.postMessage({"words": self.nlp.tokenize(e.data.text).wordCount()});
};