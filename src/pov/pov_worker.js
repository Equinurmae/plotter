importScripts('https://unpkg.com/compromise@latest/builds/compromise.min.js');
importScripts('https://unpkg.com/compendium-js@0.0.31/dist/compendium.minimal.js');

self.onmessage = function(e) {
    let doc = self.nlp(e.data.text);

    var compendium;
    
    // if compendium crashes, return empty data
    try {
        compendium = self.compendium.analyse(e.data.text);
    } catch (error) {
        compendium = [];
    }

    self.postMessage(self.count_pronouns(doc, compendium));
};

// function to count the pronouns
function count_pronouns(doc, compendium) {
    let first = doc.match("(I|me|my|mine|myself|we|us|our|ours|ourselves)").length;
    let second = doc.match("(you|your|yours|yourself|yourselves)").length;
    let third = doc.match("(he|him|his|himself|she|her|hers|herself|it|its|itself|they|them|their|theirs|themselves|themself)").length;
    let male = doc.match("(he|him|his|himself)").length;
    let female = doc.match("(she|her|hers|herself)").length;
    let neutral = doc.match("(it|its|itself|they|them|their|theirs|themselves|themself)").length;

    let entities = [];

    compendium.forEach(x => entities.push("entities" in x ? x.entities.map(e => e.raw) : []));

    return {"pronouns" : {"1st": first, "2nd": second, "3rd": third, "male": male, "female": female, "neutral": neutral}, "entities" : entities.filter(x => x != "")};
}

