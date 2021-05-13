// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

import {ma} from 'moving-averages';

const d3 = require("d3");
const regression = require("d3-regression");

// global variables

var options = {"readability": true, "words": true, "average": true, "z-scores": false, "detail": 100};

var data = [];

// web workers

var messageQueue = [];

const NUM_WORKERS = 1;
var workers_returned = 0;
const workers = [];

// spawn multiple web workers
for(let i = 0; i < NUM_WORKERS; i++) {
  workers.push(new Worker("pacing_worker.js"));

  workers[i].onmessage = function(e) {
    document.getElementById("debug").innerHTML = "Message received.";
  
    // update the data
    data.push(e.data);
  
    // check if more messages in queue
    if(messageQueue.length > 0) {
      workers[i].postMessage({"text": messageQueue.pop()});
    } else {
      workers_returned++;
  
      // check if all workers have returned
      if(workers_returned == NUM_WORKERS) {
        // update the line chart   
        data.reverse();
        draw_chart();
      }
    }
  };

}

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

// function triggers on readability option change
function onReadability() {
  options.readability = document.getElementById("readability").checked;
  refresh();
}

// function triggered on words option change
function onWords() {
  options.words = document.getElementById("words").checked;
  refresh();
}

// function triggered on average option change
function onAverage() {
  options.average = document.getElementById("average").checked;
  refresh();
}

// functino triggered on z-scores option changed
function onZ() {
  options["z-scores"] = document.getElementById("z-scores").checked;
  refresh();
}

// function triggered on average window slider changed
function onDetail() {
  options.detail = document.getElementById("detail").value;
  document.getElementById("detail_label").innerHTML = options.detail;
  refresh();
}

// function to update the spinners
function loading() {
  document.getElementById("pacing_vis").innerHTML = `<div class="ms-Spinner"></div>`;

  var SpinnerElements = document.querySelectorAll(".ms-Spinner");
  for (var i = 0; i < SpinnerElements.length; i++) {
    new fabric['Spinner'](SpinnerElements[i]);
  }
}

// main function
function refresh() {
  // reset all data
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
      
      // get selection and split into paragraphs
      if(selection.text.length == 0) {
        // no selection, so use body text
        messageQueue = paragraphs.items.map(paragraph => paragraph.text);
      } else {
        // use selection
        let results = selection.paragraphs.items.map(paragraph => paragraph.text);

        if(results.length > 1) {
          // one or more paragraphs selected
          // use regex to find the intersections
          let wholeText = results.join('\r');
          let match = new RegExp('(.*)' + selection.text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(.*)', 'g').exec(wholeText);

          if(match != null) {
            let firstParagraphMatch = new RegExp(match[1].replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(.*)', 'g').exec(results[0]);
            results[0] = firstParagraphMatch[1];

            let lastParagraphMatch = new RegExp('(.*)' + match[2].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g').exec(results[results.length-1]);
            results[results-1] = lastParagraphMatch[1];
          }
        } else {
          // one or fewer paragraphs selected
          results[0] = selection.text;
        }

        // update message queue
        messageQueue = results;
      }

        // start web worker processing
        workers_returned = 0;

        for(let i = 0; i < NUM_WORKERS; i++) {
          workers[i].postMessage({"id": i, "text": messageQueue.pop()});
        }
        
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

// function to calculate moving averages
function movingAverage(data, width) {
  return Array.from(ma(data, width));
}

// function to calculate z-scores
function z(point, mean, deviation) {
  return Math.abs((point.readability - mean) / deviation);
}

// functino to draw the line chart
function draw_chart() {
  document.getElementById("pacing_vis").innerHTML = "";

  var margin = {top: 50, right: 50, bottom: 50, left: 60}
  , width = window.innerWidth - margin.left - margin.right
  , height = window.innerHeight - margin.top - margin.bottom;

  // get chart data
  let readability = data.map(x => isNaN(x.readability) ? 0 : x.readability);
  let xs = readability.map(function(x, i) {return {"readability": x, "index": i + 1};});
  let averages = movingAverage(readability, Math.ceil(data.length * 0.15));

  let mean = d3.mean(readability);
  let deviation = d3.deviation(readability);
  let zs = data.map(d => z(d, mean, deviation));

  // init detail slider (currently hidden on front end)
  document.getElementById("detail").min = 2;
  document.getElementById("detail").max = data.length - 2;

  // init loess generator
  let loess = regression.regressionLoess()
    .y(d => d.readability)
    .x(d => d.index)
    .bandwidth(0.25);

  // create the scales for the line chart
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

  // create the line generators for the chart
  var readabilityLine = d3.line()
    .x(function(d) { return xScaleReadability(d.readability); })
    .y(function(d, i) { return yScale(i); })
    .curve(d3.curveBasis)
    .defined(d => !isNaN(d.readability));

  var averageLine = d3.line()
    .x(function(d) { return xScaleReadability(d[1]); })
    .y(function(d) { return yScale(d[0]); })
    .curve(d3.curveBasis);

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

  // draw the svg
  d3.select("svg").remove();

  var svg = d3.select("#pacing_vis").append("svg")
    .attr("width", width + margin.left + margin.right)
    .attr("height", height + margin.top + margin.bottom)
  .append("g")
    .attr("transform", "translate(" + margin.left + "," + margin.top + ")");

  // draw the y axis
  svg.append("g")
    .attr("class", "y axis")
    .call(d3.axisLeft(yScale));

  // draw the word count line
  if(options.words) {
    svg.append("path")
      .datum(data)
      .attr("class", "wordLine")
      .attr("d", wordLine);
  }

  // draw the readability line
  if(options.readability) {
    svg.append("path")
      .datum(data)
      .attr("class", "line")
      .attr("d", readabilityLine);
  }

  // draw the z-scores line
  if(options["z-scores"]) {
    svg.append("path")
      .datum(zs)
      .attr("class", "lineZ")
      .attr("d", zLine);
  }

  // draw the average line
  if(options.average) {
    svg.append("path")
      .datum(loess(xs))
      .attr("class", "lineAverage")
      .attr("d", averageLine);
  }

  // draw the x axis
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