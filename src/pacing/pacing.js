// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

import {
  ma, dma, ema, sma, wma
} from 'moving-averages';

const d3 = require("d3");

var options = {"readability": true, "words": true, "average": true, "z-scores": false, "detail": 100};

var data = [];

var messageQueue = [];

const worker = new Worker("pacing_worker.js");

worker.onmessage = function(e) {
  document.getElementById("debug").innerHTML = "Message received.";

  data.push(e.data);

  if(messageQueue.length > 0) {
    worker.postMessage({"text": messageQueue.pop()});
  } else {
    data.reverse();
    draw_chart();
  }
};

/* global document, Office, Word */

Office.onReady(info => {
  if (info.host === Office.HostType.Word) {
    // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("detail").value = options.detail;

    document.getElementById("refresh").onclick = refresh;
    document.getElementById("readability").onchange = onReadability;
    document.getElementById("words").onchange = onWords;
    document.getElementById("average").onchange = onAverage;
    document.getElementById("detail").oninput = onDetail;
    document.getElementById("z-scores").onchange = onZ;

    refresh(true);

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

function onReadability() {
  options.readability = document.getElementById("readability").checked;
  refresh();
}

function onWords() {
  options.words = document.getElementById("words").checked;
  refresh();
}

function onAverage() {
  options.average = document.getElementById("average").checked;
  refresh();
}

function onZ() {
  options["z-scores"] = document.getElementById("z-scores").checked;
  refresh();
}

function onDetail() {
  options.detail = document.getElementById("detail").value;
  document.getElementById("detail_label").innerHTML = options.detail;
  refresh();
}

function loading() {
  document.getElementById("pacing_vis").innerHTML = `<div class="ms-Spinner"></div>`;

  var SpinnerElements = document.querySelectorAll(".ms-Spinner");
  for (var i = 0; i < SpinnerElements.length; i++) {
    new fabric['Spinner'](SpinnerElements[i]);
  }
}

function refresh(first = false) {
  loading();
  
  data = [];

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

function movingAverage(data, width) {
  return Array.from(ma(data, width));
}

function z(point, mean, deviation) {
  return Math.abs((point.readability - mean) / deviation);
}

function draw_chart() {
  document.getElementById("pacing_vis").innerHTML = "";

  var margin = {top: 50, right: 50, bottom: 50, left: 60}
  , width = window.innerWidth - margin.left - margin.right
  , height = window.innerHeight - margin.top - margin.bottom;

  let readability = data.map(x => isNaN(x.readability) ? 0 : x.readability);

  document.getElementById("detail").min = 2;
  document.getElementById("detail").max = data.length - 2;

  let averages = movingAverage(readability, Math.ceil(data.length * 0.15));

  let mean = d3.mean(readability);
  let deviation = d3.deviation(readability);

  let zs = data.map(d => z(d, mean, deviation));

  var xScaleReadability = d3.scaleLinear()
    .domain([0, d3.max(readability)])
    .range([0, width]);

  var axisScale = d3.scaleOrdinal()
    .domain(["fast", "average", "slow"])
    .range([0, width/2, width]);;

  var xScaleWords = d3.scaleLinear()
    .domain([0, d3.max(data.map(x => x.words))])
    .range([0, width]);

  var xScaleZ = d3.scaleLinear()
    .domain([0, d3.max(zs)])
    .range([0, width]);

  var yScale = d3.scaleLinear()
    .domain([data.length-1,0])
    .range([height, 0]);

  var readabilityLine = d3.line()
    .x(function(d) { return xScaleReadability(d.readability); })
    .y(function(d, i) { return yScale(i); })
    .curve(d3.curveBasis)
    .defined(d => !isNaN(d.readability));

  var averageLine = d3.line()
    .x(function(d) { return xScaleReadability(d); })
    .y(function(d, i) { return yScale(i); })
    .curve(d3.curveBasis)
    .defined(d => d != undefined);

  var wordLine = d3.line()
    .x(function(d) { return xScaleWords(d.words); })
    .y(function(d, i) { return yScale(i); })
    .curve(d3.curveMonotoneX)
    .defined(d => d.words != undefined);

  var zLine = d3.line()
    .x(function(d) { return xScaleZ(d); })
    .y(function(d, i) { return yScale(i); })
    .curve(d3.curveBasis)
    .defined(d => d != undefined);

  d3.select("svg").remove();

  var svg = d3.select("#pacing_vis").append("svg")
    .attr("width", width + margin.left + margin.right)
    .attr("height", height + margin.top + margin.bottom)
  .append("g")
    .attr("transform", "translate(" + margin.left + "," + margin.top + ")");

  svg.append("g")
    .attr("class", "y axis")
    .call(d3.axisLeft(yScale));

  if(options.words) {
    svg.append("path")
      .datum(data)
      .attr("class", "wordLine")
      .attr("d", wordLine);
  }

  if(options.readability) {
    svg.append("path")
      .datum(data)
      .attr("class", "line")
      .attr("d", readabilityLine);
  }

  if(options["z-scores"]) {
    svg.append("path")
      .datum(zs)
      .attr("class", "lineZ")
      .attr("d", zLine);
  }

  if(options.average) {
    svg.append("path")
      .datum(averages)
      .attr("class", "lineAverage")
      .attr("d", averageLine);
  }

  svg.append("g")
    .attr("class", "x axis")
    .attr("transform", "translate(0,0)")
    .call(d3.axisTop(axisScale));

  svg.append("text")             
    .attr("transform",
        "translate(" + (width/2) + " ," + 
                       (- 30) + ")")
    .style("text-anchor", "middle")
    .text("Pacing");

  svg.append("text")
      .attr("transform", "rotate(-90)")
      .attr("y", 0 - margin.left + 10)
      .attr("x",0 - (height / 2))
      .attr("dy", "1em")
      .style("text-anchor", "middle")
      .text("Paragraph #"); 

}