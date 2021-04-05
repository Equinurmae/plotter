// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

const d3 = require("d3");

var messageQueue = [];

const worker = new Worker("pov_worker.js");

worker.onmessage = function(e) {
  document.getElementById("debug").innerHTML = "Message received.";

  updatePronounData(e.data.pronouns);
  redrawPieChart(messageQueue.length == 0);

  if(messageQueue.length > 0) {  
    worker.postMessage({"text": messageQueue.pop()});
  } else {
    document.getElementById("guess").innerHTML = Object.entries(pronoun_data).sort((a,b) => a[1] - b[1])[2][0] + " person";
  }
};

var margin = {top: 30, right: 30, bottom: 30, left: 30}
, width = window.innerWidth - margin.left - margin.right
, height = window.innerWidth - margin.top - margin.bottom;

var radius = Math.min(width, height) / 2 - margin.top;

var svg = d3.select("#pie_chart_vis")
  .append("svg")
  .attr("width", width + margin.left + margin.right)
  .attr("height", height + margin.top + margin.bottom)
  .append("g")
  .attr("transform",
        "translate(" + ((width / 2) + margin.left) + "," + ((height / 2) + margin.top) + ")");

var pronoun_data = {"1st": 0, "2nd": 0, "3rd": 0};

var colourScale = d3.scaleOrdinal(d3.schemeSet3);

colourScale.domain(Object.keys(pronoun_data));

var pie = d3.pie()
  .value((d) => d.value);

var arcGenerator = d3.arc()
  .innerRadius(0)
  .outerRadius(radius);

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.

    document.getElementById("refresh").onclick = refresh;

    refresh();
    drawPieChart();

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

function refresh() {
  resetPronounData();

  Word.run(function (context) {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");

    var selection = context.document.getSelection();

    selection.load("text");

    selection.paragraphs.load("text");

    return context.sync()
      .then(function() {
        document.getElementById("debug").innerHTML = "Message sending...";
        
        if(selection.text.length == 0) {
          messageQueue = paragraphs.items.map(paragraph => paragraph.text);
        } else {
          let results = selection.paragraphs.items.map(paragraph => paragraph.text);

          if(results.length > 1) {
            let wholeText = results.join('\r');
            let match = new RegExp('(.*)' + selection.text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(.*)', 'g').exec(wholeText);

            if(match != null) {
              let firstParagraphMatch = new RegExp(match[1].replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(.*)', 'g').exec(results[0]);
              results[0] = firstParagraphMatch[1];

              let lastParagraphMatch = new RegExp('(.*)' + match[2].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g').exec(results[results.length-1]);
              results[results-1] = lastParagraphMatch[1];
            }
          } else {
            results[0] = selection.text;
          }

          messageQueue = results;
        }
        
        worker.postMessage({"text": messageQueue.pop()});
        document.getElementById("debug").innerHTML = "Message sent.";
      })
      .then(context.sync);
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

function drawPieChart() {
  var data_ready = pie(Object.entries(pronoun_data).map(function(x) {return {"key": x[0], "value": x[1]}; }));

  svg
  .selectAll('mySlices')
  .data(data_ready)
  .enter()
  .append('path')
  .attr("class", "arc")
  .attr('d', arcGenerator)
  .attr('fill', (d) => colourScale(d.data.key))
  .attr("stroke", "white")
  .style("stroke-width", "2px");

svg
  .selectAll('mySlices')
  .data(data_ready)
  .enter()
  .append('text')
  .attr("class", "label")
  .text(function(d){ return d.data.key})
  .attr("transform", function(d) { return "translate(" + arcGenerator.centroid(d) + ")";  })
  .style("text-anchor", "middle")
  .style("font-size", 17)
  .style("opacity", d => d.data.value > 0 ? 100 : 0);
}

function redrawPieChart(animate) {
  var data_ready = pie(Object.entries(pronoun_data).map(function(x) {return {"key": x[0], "value": x[1]}; }));

  svg
    .selectAll('.arc')
    .data(data_ready)
    .transition().duration(animate ? 800 : 1)
    .attr('d', arcGenerator)
    .attr('fill', (d) => colourScale(d.data.key))
    .attr("stroke", "white")
    .style("stroke-width", "2px");

  svg
    .selectAll('.label')
    .data(data_ready)
    .transition().duration(animate ? 800 : 10)
    .text(function(d){ return d.data.key})
    .attr("transform", function(d) { return "translate(" + arcGenerator.centroid(d) + ")";  })
    .style("text-anchor", "middle")
    .style("font-size", 17)
    .style("opacity", d => d.data.value > 0 ? 100 : 0);
}

function updatePronounData(newData) {
  pronoun_data["1st"] += newData["1st"];
  pronoun_data["2nd"] += newData["2nd"];
  pronoun_data["3rd"] += newData["3rd"];
}

function resetPronounData() {
  pronoun_data = {"1st": 0, "2nd": 0, "3rd": 0};
}