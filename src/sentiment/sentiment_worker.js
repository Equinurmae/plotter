importScripts('https://unpkg.com/compendium-js@0.0.31/dist/compendium.minimal.js');

self.onmessage = function(e) {
    let doc = self.compendium.analyse(e.data.text);

    self.postMessage(self.polarity(doc));
};

function polarity(doc) {
    let polarity = 0;
    doc.forEach(x => polarity += ("profile" in x ? x.profile.sentiment : 0));

    return {"polarity": polarity};
}