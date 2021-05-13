// images references in the manifest
import "../../assets/icon-16.png";
import "../../assets/icon-32.png";
import "../../assets/icon-80.png";

import {ma} from 'moving-averages';

const d3 = require("d3");

// global variables

var polarity = [];

// web workers

var messageQueue = [];

const worker = new Worker("sentiment_worker.js");

worker.onmessage = function(e) {
  document.getElementById("debug").innerHTML = "Message received.";

  // update polarity data
  polarity.push(e.data.polarity);

  let avg = polarity.reduce((a, b) => a + b, 0) / polarity.length;

  // check if messages still left in queue
  if(messageQueue.length > 0) { 
    worker.postMessage({"text": messageQueue.pop().replace(/[^\x20-\x7E]/g, '')});
  } else {
    // update polarity message bar
    document.getElementById("polarity").innerHTML = `<div class="ms-MessageBar ms-MessageBar--` + (avg == 0 ? 'warning' : (avg < 0 ? 'error' : 'success')) + `">
      <div class="ms-MessageBar-content">
        <div class="ms-MessageBar-icon">
          <i class="ms-Icon ms-Icon--Completed"></i>
        </div>
        <div class="ms-MessageBar-text">
          The polarity of this text is <span style="font-weight: 700">` + Math.abs(avg.toFixed(2) * 100) + '% ' + (avg == 0 ? 'neutral' : (avg < 0 ? 'negative' : 'positive')) + `</span>.
        </div>
      </div>
    </div>`;
  
    // draw charts
    draw_chart();

    polarity.reverse();
    drawLineChart();
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
    document.getElementById("refresh").onclick = refresh;

    refresh();

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

// function to update spinners
function loading() {
  document.getElementById("polarity").innerHTML = `<div class="ms-Spinner"></div>`;
  document.getElementById("sentiment_vis").innerHTML = `<br><div class="ms-Spinner"></div>`;
  document.getElementById("line_chart_vis").innerHTML = `<div class="ms-Spinner"></div>`;

  var SpinnerElements = document.querySelectorAll(".ms-Spinner");
  for (var i = 0; i < SpinnerElements.length; i++) {
    new fabric['Spinner'](SpinnerElements[i]);
  }
}

// main function
function refresh() {
  loading();
  polarity = [];

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
        worker.postMessage({"text": messageQueue.pop().replace(/[^\x20-\x7E]/g, '')});
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

// function to draw the dummy spider chart
// code adapted from https://yangdanny97.github.io/blog/2019/03/01/D3-Spider-Chart
function draw_chart() {
  document.getElementById("sentiment_vis").innerHTML = "";
  d3.select("svg").remove();

  var margin = {top: 50, right: 50, bottom: 50, left: 50}
  , width = window.innerWidth - margin.left - margin.right
  , height = window.innerWidth - margin.top - margin.bottom;

  var svg = d3.select("#sentiment_vis")
    .append("svg")
    .attr("width", width + margin.left + margin.right)
    .attr("height", height + margin.top + margin.bottom)
    .append("g")
    .attr("transform",
          "translate(" + margin.left + "," + margin.top + ")");
  
  // create the scales for the chart
  let radialScale = d3.scaleLinear()
    .domain([0,10])
    .range([0,width/2]);

  let ticks = [2,4,6,8,10];

  ticks.forEach(t =>
    svg.append("circle")
    .attr("cx", width/2)
    .attr("cy", width/2)
    .attr("fill", "none")
    .attr("stroke", "gray")
    .attr("r", radialScale(t))
  );

  // dummy chart data
  let features = ["Joy", "Sadness", "Anger", "Fear", "Disgust"];
  let data = [{"Joy": Math.random() * 10, "Sadness": Math.random() * 10, "Anger": Math.random() * 10, "Fear": Math.random() * 10, "Disgust": Math.random() * 10}];

  // helper functions
  function angleToCoordinate(angle, value){
    let x = Math.cos(angle) * radialScale(value);
    let y = Math.sin(angle) * radialScale(value);
    return {"x": width/2 + x, "y": width/2 - y};
  }

  // draw the chart
  for (var i = 0; i < features.length; i++) {
    let ft_name = features[i];
    let angle = (Math.PI / 2) + (2 * Math.PI * i / features.length);
    let line_coordinate = angleToCoordinate(angle, 10);
    let label_coordinate = angleToCoordinate(angle, 10.5);

    //draw axis line
    svg.append("line")
    .attr("x1", width/2)
    .attr("y1", width/2)
    .attr("x2", line_coordinate.x)
    .attr("y2", line_coordinate.y)
    .attr("stroke","black");

    //draw axis label
    svg.append("text")
    .attr("x", label_coordinate.x)
    .attr("y", label_coordinate.y)
    .text(ft_name);

    let line = d3.line()
      .x(d => d.x)
      .y(d => d.y);

    // calculate path coordinates
    function getPathCoordinates(data){
      let coordinates = [];
      for (var i = 0; i < features.length; i++){
          let ft_name = features[i];
          let angle = (Math.PI / 2) + (2 * Math.PI * i / features.length);
          coordinates.push(angleToCoordinate(angle, data[ft_name]));
      }

      coordinates.push(coordinates[0])
      return coordinates;
    }

    let coordinates = getPathCoordinates(data[0]);

    svg.append("path")
      .datum(coordinates)
      .attr("d",line)
      .attr("stroke-width", 3)
      .attr("stroke", "#0084ff")
      .attr("fill", "#accdeb")
      .attr("stroke-opacity", 1)
      .attr("opacity", 0.5);
  }
}

// function to calculate the moving averages
function movingAverage(data, width) {
  return Array.from(ma(data, width));
}

// function to draw the line chart
function drawLineChart() {
  document.getElementById("line_chart_vis").innerHTML = "";

  var line_margin = {top: 50, right: 50, bottom: 50, left: 60}
    , line_width = window.innerWidth - line_margin.left - line_margin.right
    , line_height = window.innerHeight - line_margin.top - line_margin.bottom;

  var averages = movingAverage(polarity, Math.ceil(polarity.length * 0.15))

  // create the scales for the chart
  var xScaleLine = d3.scaleLinear()
    .domain([-1, 1])
    .range([0, line_width]);

  var yScaleLine = d3.scaleLinear()
      .domain([polarity.length-1,0])
      .range([line_height, 0]);

  // create the line generator
  var line = d3.line()
    .x(function(d) { return xScaleLine(d); })
    .y(function(d, i) { return yScaleLine(i); })
    .curve(d3.curveBasis)
    .defined(d => d != undefined);

  // draw the svg
  d3.select("#line_chart_vis").select("svg").remove();

  var line_svg = d3.select("#line_chart_vis").append("svg")
    .attr("width", line_width + line_margin.left + line_margin.right)
    .attr("height", line_height + line_margin.top + line_margin.bottom)
    .append("g")
      .attr("transform", "translate(" + line_margin.left + "," + line_margin.top + ")");

  // draw the y axis
  line_svg.append("g")
      .attr("class", "y-axis-line")
      .attr("transform", "translate(" + (line_width / 2) + ",0)")
      .call(d3.axisLeft(yScaleLine));

  // draw the lines
  line_svg.append("path")
      .datum(polarity)
      .attr("class", "line")
      .attr("stroke", "#0084ff")
      .attr("fill", "none")
      .attr("d", line);

  line_svg.append("path")
    .datum(averages)
    .attr("class", "lineAverage")
    .attr("d", line);

  // draw the x axis
  line_svg.append("g")
    .attr("class", "x-axis-line")
    .attr("transform", "translate(0,0)")
    .call(d3.axisTop(xScaleLine).ticks(5));

  line_svg.append("text")             
  .attr("transform",
        "translate(" + (line_width/2) + " ," + 
                       (- 30) + ")")
  .style("text-anchor", "middle")
  .text("Polarity");

  line_svg.append("text")
      .attr("transform", "rotate(-90)")
      .attr("y", (line_width / 2) - line_margin.left + 10)
      .attr("x",0 - (line_height / 2))
      .attr("dy", "1em")
      .style("text-anchor", "middle")
      .text("Paragraph #");
}