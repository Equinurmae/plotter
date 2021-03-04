importScripts('https://unpkg.com/compromise@latest/builds/compromise.min.js');

self.onmessage = function(e) {
    self.postMessage({"words": self.nlp.tokenize(e.data.text).wordCount()});
};